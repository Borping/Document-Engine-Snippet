# Document-Engine-Snippet
A small demo of a document engine that programmatically edits a template document and saves/uploads it with an appropriate naming scheme.

## Features  
- **Document Automation**: Programmatically populates a template document with student data, log entries, and summaries into tables.
- **Cloud Integration**: Secure uploading of generated documents to Google Drive with API authentication.
- **Data Extraction and Analysis**: Demo component for parsing the generated document for accommodation usage to inform future decisions.
- **File Management**: Uses a robust file naming convention including student name, date, subject, and current user.

### Prerequisites  
- Python 3.11.9+ 

## Requirements  

```
logging
datetime
pathlib
sys
typing
```
- As you can immediately tell, this project is missing a sizeable amount of dependencies, as it's part of a much larger proprietary project. It is meant for demonstration purposes only.

### Visuals
![Template Before Demo](https://i.imgur.com/qqPGs5M.png)
![Template After Demo](https://i.imgur.com/KLiz436.png)
![Output Document Analysis](https://i.imgur.com/Ayk98QM.png)
