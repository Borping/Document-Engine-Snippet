# run_doc_engine_demo.py
import logging
import datetime
from pathlib import Path
import sys
from typing import Tuple, Optional # type hints

project_root = Path(__file__).resolve().parent
sys.path.insert(0, str(project_root))

from ez_iep.core.config import ConfigManager
from ez_iep.drive.gdrive_service import GoogleDriveService
from ez_iep.drive.drive_io_manager import DriveIOManager
from ez_iep.doc_engine.document_engine import DocumentProcessor, generate_document_filename, analyze_document_accommodation_usage
from ez_iep.core.constants import EZ_IEP_SHARED_FOLDER_NAME

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("__main__")

def check_text_in_doc(element_container, text_to_find: str) -> bool:
    """Recursively checks if specific text is present in paragraphs of the element_container (doc or cell)."""
    for para in element_container.paragraphs:
        if text_to_find in para.text:
            return True
    for table in element_container.tables: # iterate tables in container
        for row in table.rows:
            for cell in row.cells:
                if check_text_in_doc(cell, text_to_find): # recursive call for each cell
                    return True
    return False

def run_demo():
    logger.info("Starting Demo")

    config_mgr = ConfigManager()
    gdrive_service = GoogleDriveService()

    if not gdrive_service.is_authenticated():
        logger.info("Google Drive Service not authenticated. Attempting authentication...")
        gdrive_service.authenticate_and_initialize()
        if not gdrive_service.is_authenticated():
            logger.error("Google Drive authentication failed. Exiting demo.")
            return
    else:
        logger.info("Google Drive Service already authenticated.")

    drive_io_mgr = DriveIOManager(gdrive_service, config_mgr)

    if not drive_io_mgr.ensure_shared_folder_is_set(attempt_find_or_create=True):
        logger.error(f"Could not ensure shared folder '{EZ_IEP_SHARED_FOLDER_NAME}' is set. Exiting.")
        return
    logger.info(f"Using shared Drive folder ID: {drive_io_mgr.get_shared_folder_id()}")

    template_file_name = "EZ-IEP Template Example - No Drawing - v1.1.docx" # latest template
    template_path = project_root / "templates" / template_file_name
    
    if not template_path.exists():
        logger.error(f"Template file not found at: {template_path}")
        return

    doc_processor = DocumentProcessor(str(template_path))
    logger.info("--- [DEBUG DUMP] Document content after initial load: ---")
    found_accommodations_label = False
    found_pref_seating = False
    for i, para in enumerate(doc_processor.doc.paragraphs):
        if "Accommodations:" in para.text or "Preferential Seating" in para.text or "{inc_teacher}" in para.text:
            logger.info(f"[DEBUG DUMP] Paragraph {i} TEXT: '{para.text}'")
        if "Accommodations:" in para.text:
            found_accommodations_label = True
        if "A. Preferential Seating" in para.text:
            found_pref_seating = True

    for t_idx, table in enumerate(doc_processor.doc.tables):
        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                for i, para in enumerate(cell.paragraphs):
                    if "Accommodations:" in para.text or "A. Preferential Seating" in para.text:
                        logger.info(f"[DEBUG DUMP] Table {t_idx}, Row {r_idx}, Cell {c_idx}, Para {i} TEXT: '{para.text}'")
                    if "Accommodations:" in para.text:
                        found_accommodations_label = True
                    if "A. Preferential Seating" in para.text:
                        found_pref_seating = True

    header_data = {
        "student": "Jane Doe",
        "weeks_of": "May 12 - May 16, 2025",    # for page 1
        "weeks_of2": "May 19 - May 23, 2025",   # for page 2
        "inc_minutes": "120",
        "subject": "Math",
        "inc_teacher": "Dr. Greg"
    }
    current_user = "Test User"
    log_identifier_for_filename = "LogV6"
    log_date = datetime.date(2025, 5, 12)

    doc_processor.fill_document_header(
        student=header_data["student"],
        weeks_of=header_data["weeks_of"],
        inc_minutes=header_data["inc_minutes"],
        subject=header_data["subject"],
        inc_teacher=header_data["inc_teacher"],
        weeks_of2=header_data["weeks_of2"] # pass our new data
    )

    # fill daily log
    doc_processor.update_daily_log_entry(
        week_num=1, day_abbrev="mon", used_accommodations=['A', 'C'],
        notes="Started strong.", service_minutes="30"
    )

    doc_processor.update_daily_log_entry(
        week_num=1, day_abbrev="tues", used_accommodations=['B'],
        notes=None, service_minutes="30" # empty note -> space
    )
    doc_processor.update_daily_log_entry(
        week_num=1, day_abbrev="wed", used_accommodations=['A', 'D', 'E'],
        notes="I think the student did very well today.", service_minutes="30"
    )
    doc_processor.fill_weekly_total(week_num=1, total_minutes_for_week="90")

    # weeks 3 and 4 for page 2
    doc_processor.update_daily_log_entry(
        week_num=3, day_abbrev="mon", used_accommodations=['A', 'F'],
        notes="Week 3 start.", service_minutes="30"
    )
    doc_processor.update_daily_log_entry(
        week_num=4, day_abbrev="fri", used_accommodations=['A', 'G', 'H'],
        notes="End of 4th week.", service_minutes="30"
    )
    doc_processor.fill_weekly_total(week_num=3, total_minutes_for_week="30")
    doc_processor.fill_weekly_total(week_num=4, total_minutes_for_week="30")

    output_filename_str = generate_document_filename(
        date_obj=log_date, user=current_user, student_name=header_data["student"],
        student_id="JD_S456", subject=header_data["subject"],
        log_identifier=log_identifier_for_filename
    )

    local_save_dir = project_root / "generated_docs"
    if not local_save_dir.exists():
        local_save_dir.mkdir(parents=True, exist_ok=True)
    local_output_path = local_save_dir / output_filename_str
    doc_processor.save(str(local_output_path))
    logger.info(f"Document saved locally to: {local_output_path}")

    try:
        document_bytes = doc_processor.get_document_bytes()
        logger.info(f"Attempting to upload '{output_filename_str}' to Google Drive...")
        
        upload_result = drive_io_mgr.upload_encrypted_file( # Assuming this is the correct method name now
            file_name=output_filename_str,
            plaintext_content=document_bytes,
            mime_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

        if upload_result:
            if isinstance(upload_result, tuple) and len(upload_result) == 2:
                file_id, web_view_link = upload_result
                logger.info(f"File uploaded successfully to Drive! File ID: {file_id}, Link: {web_view_link}")
            elif isinstance(upload_result, str):
                file_id = upload_result
                logger.info(f"File uploaded successfully to Drive! File ID: {file_id}")
            else:
                logger.info(f"File upload to Drive completed. Result: {upload_result}")
        else:
            logger.error("File upload to Drive failed.")

    except AttributeError as ae:
        logger.error(f"AttributeError during Drive upload: {ae}.", exc_info=True)
    except TypeError as te:
        logger.error(f"TypeError during Drive upload: {te}.", exc_info=True)
    except Exception as e:
        logger.error(f"General error during Drive upload process: {e}", exc_info=True)

    # analyze the generated output document
    if local_output_path.exists():
        logger.info(f"\n--- Analyzing accommodations in: {local_output_path} ---")
        # for a 4-week template with 5 days each = 20 potential service days
        # analysis will count actual filled entries.
        # total_service_days should ideally account for absences
        num_potential_days = 20 
        usage_report = analyze_document_accommodation_usage(str(local_output_path), total_service_days=num_potential_days)
        
        if usage_report:
            print("\n--- Accommodation Usage Report ---")
            for letter, data in usage_report.items():
                # print if count > 0 or if it was defined (to be concise)
                if data['count'] > 0 or letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ": # ensure all acc. from template are shown
                     print(f"{letter}. {data.get('description', 'N/A')}: {data['count']}/{data['total_days']}")
            print("---------------------------------")
        else:
            logger.warning("Accommodation usage analysis did not produce a report.")


    logger.info("Demo Finished")

if __name__ == "__main__":
    run_demo()