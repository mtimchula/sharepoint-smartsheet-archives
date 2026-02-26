import msal
import os
from dotenv import load_dotenv
import requests
import pandas as pd
import io
from datetime import datetime
import os
import logging

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

load_dotenv()


class SharePoint:
    def __init__(self, site_name):
        self.tenant_id = os.getenv("TENANT_ID")
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")

        self.base_url = "bhaz.sharepoint.com:/sites"
        # self.site_name = site_name
        self.site_url = f"{self.base_url}/{site_name}"

        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"

        self.access_token = self._get_token()

    def _get_token(self):
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret,
        )
        token_result = app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )

        if "access_token" not in token_result:
            raise Exception(token_result)

        return token_result["access_token"]

    def get_site_id(self):
        headers = {
            "Authorization": f"Bearer {self.access_token}",
        }

        site_resp = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{self.site_url}",
            headers=headers,
        )

        site_resp.raise_for_status()
        return site_resp.json()["id"]

    def download_file(self, file_path):

        headers = {
            "Authorization": f"Bearer {self.access_token}",
        }

        site_id = self.get_site_id()
        download_url = (
            f"https://graph.microsoft.com/v1.0/"
            f"sites/{site_id}/drive/root:/{file_path}:/content"
        )

        response = requests.get(download_url, headers=headers)
        response.raise_for_status()

        return response.content

    def upload_file(self, content, file_path):

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "text/csv",
        }
        site_id = self.get_site_id()

        upload_url = (
            f"https://graph.microsoft.com/v1.0/"
            f"sites/{site_id}/drive/root:/{file_path}:/content"
        )

        response = requests.put(
            upload_url,
            headers=headers,
            data=content.encode("utf-8"),
        )

        response.raise_for_status()

    def list_files_in_folder(self, folder_path: str) -> list[dict]:
        headers = {
            "Authorization": f"Bearer {self.access_token}",
        }
        site_id = self.get_site_id()

        url = (
            f"https://graph.microsoft.com/v1.0/"
            f"sites/{site_id}/drive/root:/{folder_path}:/children"
        )

        resp = requests.get(url, headers=headers)
        resp.raise_for_status()

        # Return only files (not subfolders)
        return [item for item in resp.json()["value"] if "file" in item]

    def move_file_by_id(self, file_id: str, destination_folder_path: str):
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
        }
        site_id = self.get_site_id()

        # Resolve destination folder ID
        dest_folder_url = (
            f"https://graph.microsoft.com/v1.0/"
            f"sites/{site_id}/drive/root:/{destination_folder_path}"
        )

        dest_resp = requests.get(dest_folder_url, headers=headers)
        dest_resp.raise_for_status()
        dest_folder_id = dest_resp.json()["id"]

        # Move file
        move_url = (
            f"https://graph.microsoft.com/v1.0/"
            f"sites/{site_id}/drive/items/{file_id}"
        )

        payload = {"parentReference": {"id": dest_folder_id}}

        resp = requests.patch(
            move_url,
            headers=headers,
            json=payload,
        )
        resp.raise_for_status()


def create_csv_from_bytes(content):
    excel_buffer = io.BytesIO(content)

    df = pd.read_excel(
        excel_buffer, skiprows=1, sheet_name="Skill Definitions", engine="openpyxl"
    )
    df = df.dropna(subset=["WD Skill Name"])
    cols = ["WD Skill Name", "Definition"]

    df = df[cols]

    return df.to_csv(index=False)


def main():
    today = datetime.now().strftime("%m%d%Y")
    current_year = datetime.now().year

    sftp_folder = "SFTP - Files"

    upload_file_name = f"SFTP - Files/Skills Definitions {today}.csv"
    download_path = "General/Belong Site/BH Mercer Skills Library Definitions.xlsx"
    sftp_path = f"General/Belong Site/Skills Definition Uploads"
    upload_path = f"{sftp_path}/{upload_file_name}"
    move_path = f"{sftp_path}/{current_year}"

    latest_file_path = f"{sftp_path}/{sftp_folder}"

    conn = SharePoint(site_name="CompensationTeam9")
    xlsx_content = conn.download_file(file_path=download_path)
    csv_content = create_csv_from_bytes(xlsx_content)

    # move the old file(s) from the folder
    files = conn.list_files_in_folder(folder_path=latest_file_path)
    if files:
        latest_file = max(files, key=lambda f: f["lastModifiedDateTime"])
        conn.move_file_by_id(
            file_id=latest_file["id"],
            destination_folder_path=move_path,
        )
    else:
        logger.info("No existing files found; skipping move.")

    # # upload the newest file
    conn.upload_file(content=csv_content, file_path=upload_path)


if __name__ == "__main__":
    main()
