# SharePoint-Smartsheet Data Integration
This Python script facilitates the integration of data between SharePoint and Smartsheet platforms. It enables users to perform tasks such as reading files from SharePoint, retrieving data from Smartsheet, combining dataframes, uploading files to SharePoint, and deleting existing data from Smartsheet.

# Requirements
Python 3.x
Required Python libraries:
- office365
- pandas
- smartsheet

# Installation
- Clone this repository to your local machine.
- Install the required Python libraries using pip:
```
pip install -r requirements.txt
```
# Usage
Ensure you have valid credentials for accessing both SharePoint and Smartsheet.
Modify the creds.py file with your SharePoint and Smartsheet credentials.
Run the script by executing the following command:
```
python main.py
```
# Functionality
- login_sharepoint(): Authenticates the user with SharePoint.
- read_sharepoint_file(login_sp, sp_file): Reads files from SharePoint and returns them as a list of Pandas DataFrames.
- get_smartsheet_data(smartsheet_id): Retrieves data from Smartsheet and returns it as a list of DataFrames.
- combine_dataframes(sharepoint_dfs, smartsheet_dfs): Combines DataFrames from SharePoint and Smartsheet.
- upload_files_to_sp(ctx, combined_dfs, file_names): Uploads files to SharePoint.
- delete_existing_data(sheet_ids, chunk_interval=300): Deletes existing data from Smartsheet.
