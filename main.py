from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import io
import pandas as pd
import smartsheet
import time
from location import location_data
import creds


def login_sharepoint():
    """Authenticate with O365"""
    ctx_auth = AuthenticationContext(creds.BASE_URL)
    if ctx_auth.acquire_token_for_user(creds.USERNAME, creds.PASSWORD):
        ctx = ClientContext(creds.BASE_URL, ctx_auth)
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        print("Authentication successful")
        return ctx


def read_sharepoint_file(login_sp, sp_file):
    """Read each file in SharePoint folder and return them as a list of dataframes"""
    dfs = []
    for name in sp_file:  # loop through each file name and create a list of dataframes from SP
        file_name = name['name']
        response = File.open_binary(login_sp, f"{creds.RELATIVE_URL}{file_name}")
        bytes_file_obj = io.BytesIO()
        bytes_file_obj.write(response.content)
        bytes_file_obj.seek(0)  # set file object to start
        df = pd.read_csv(bytes_file_obj)
        dfs.append(df)
    return dfs


def get_smartsheet_data(smartsheet_id):
    """Get data from Smartsheet Archive file"""
    dfs = []
    for sheet in smartsheet_id:
        sheet_id = sheet['sheet_id']
        client = smartsheet.Smartsheet(creds.TOKEN)
        sheet_dict = client.Sheets.get_sheet(sheet_id).to_dict()  # turn json response into dict
        rows = sheet_dict['rows']
        data = [[cell.get('value', '') for cell in row.get('cells', [])] for row in rows]
        sheet_df = pd.DataFrame(data)
        dfs.append(sheet_df)  # append the smartsheet sheet data to the dfs dataframe list
        time.sleep(1)  # sleeping the request to limit responses to server
    return dfs


def combine_dataframes(sharepoint_dfs, smartsheet_dfs):
    """Reset the indexes of dataframes and joins them together"""
    combined_dfs = []
    for sharepoint_df, smartsheet_df in zip(sharepoint_dfs, smartsheet_dfs):  # loop through each dataframe in each list
        sharepoint_df.columns = smartsheet_df.columns  # align the columns for rach dataframe to match
        sharepoint_df.reset_index(drop=True, inplace=True)
        smartsheet_df.reset_index(drop=True, inplace=True)
        combined_df = pd.concat([sharepoint_df, smartsheet_df], ignore_index=True)
        combined_dfs.append(combined_df)
    return combined_dfs


def upload_files_to_sp(ctx, combined_dfs, file_names):
    for files, combined_df in zip(file_names, combined_dfs):  # loop through each excel name and dataframe to upload
        file_name = files['name']
        target_folder = ctx.web.get_folder_by_server_relative_url(creds.RELATIVE_URL)
        buffer = io.BytesIO()
        combined_df.to_csv(buffer, index=False)
        buffer.seek(0)
        file_content = buffer.read()
        target_folder.upload_file(file_name, file_content).execute_query()  # upload files to location in SP
    print("Files were uploaded successfully.")


def delete_existing_data(sheet_ids, chunk_interval=300):
    """Delete the existing data in each row for each sheet_id"""
    for sheet in sheet_ids:
        sheet_id = sheet['sheet_id']  # loop through the dictionary to get the sheet_ids
        client = smartsheet.Smartsheet(creds.TOKEN)  # initialize the smartsheet client
        sheet = client.Sheets.get_sheet(sheet_id)
        rows_to_delete = [row.id for row in sheet.rows]  # delete data from smartsheet
        for x in range(0, len(rows_to_delete), chunk_interval):  # delete the data in chunks
            client.Sheets.delete_rows(sheet.id, rows_to_delete[x:x + chunk_interval])
    print("Sheet data deleted")


def main():
    ctx = login_sharepoint()  # login to SP
    file_names = location_data  # access dictionary
    if ctx:
        sharepoint_df = read_sharepoint_file(ctx, file_names)
        smartsheet_df = get_smartsheet_data(file_names)
        combined_df = combine_dataframes(sharepoint_df, smartsheet_df)
        upload_files_to_sp(ctx, combined_df, file_names)
        delete_existing_data(file_names)
    else:
        print("Process failed.")


main()
