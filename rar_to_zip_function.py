import os
import shutil
import requests
import rarfile
import zipfile
import tempfile
import azure.functions as func
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# SharePoint and Azure AD Credentials (Use Azure Key Vault for security)
SP_SITE_URL = "https://yourtenant.sharepoint.com/sites/YourSite"
SP_DOC_LIBRARY = "Shared Documents"
CLIENT_ID = "your-client-id"
CLIENT_SECRET = "your-client-secret"

# Authenticate with SharePoint using Client Credentials
def get_sharepoint_context():
    ctx = ClientContext(SP_SITE_URL).with_credentials(ClientCredential(CLIENT_ID, CLIENT_SECRET))
    return ctx

def download_file_from_sharepoint(file_url, local_path):
    """ Downloads a file from SharePoint using Microsoft Graph API. """
    ctx = get_sharepoint_context()
    file_name = os.path.basename(file_url)
    file = ctx.web.get_file_by_server_relative_url(f"{SP_DOC_LIBRARY}/{file_name}")
    with open(local_path, 'wb') as f:
        f.write(file.download().execute_query())
    return True

def upload_file_to_sharepoint(local_path, file_name):
    """ Uploads a file to SharePoint using Microsoft Graph API. """
    ctx = get_sharepoint_context()
    folder = ctx.web.get_folder_by_server_relative_url(SP_DOC_LIBRARY)
    with open(local_path, 'rb') as f:
        folder.upload_file(file_name, f).execute_query()
    return True

def convert_rar_to_zip(rar_path, zip_path):
    """ Extracts RAR file and compresses it into ZIP. """
    with tempfile.TemporaryDirectory() as temp_dir:
        with rarfile.RarFile(rar_path) as rf:
            rf.extractall(temp_dir)
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for foldername, subfolders, filenames in os.walk(temp_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zf.write(file_path, arcname)

def main(req: func.HttpRequest) -> func.HttpResponse:
    try:
        file_url = req.params.get("file_url")
        if not file_url:
            return func.HttpResponse("Please provide a SharePoint file URL", status_code=400)

        rar_filename = os.path.basename(file_url)
        zip_filename = rar_filename.replace(".rar", ".zip")

        # Define local paths
        temp_dir = tempfile.gettempdir()
        rar_path = os.path.join(temp_dir, rar_filename)
        zip_path = os.path.join(temp_dir, zip_filename)

        # Step 1: Download RAR from SharePoint
        if not download_file_from_sharepoint(file_url, rar_path):
            return func.HttpResponse("Failed to download .rar file from SharePoint.", status_code=500)

        # Step 2: Convert RAR to ZIP
        convert_rar_to_zip(rar_path, zip_path)

        # Step 3: Upload ZIP back to SharePoint
        if not upload_file_to_sharepoint(zip_path, zip_filename):
            return func.HttpResponse("Failed to upload .zip file to SharePoint.", status_code=500)

        return func.HttpResponse(f"RAR converted to ZIP successfully! File: {zip_filename}", status_code=200)
    
    except Exception as e:
        return func.HttpResponse(f"Error: {str(e)}", status_code=500)
