"""
Google Drive to Windows Folder File Transfer Script

This script:
1. Connects to Google Drive using service account
2. Lists files in the specified Google Drive folder
3. Downloads files to a Windows folder
4. Optionally deletes files from Google Drive after download (move mode)
"""

import os
import logging
from typing import Optional, List, Dict, Any
from io import BytesIO

from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# Load environment variables from .env file
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class Config:
    """Configuration settings"""
    # Google Drive settings
    SERVICE_ACCOUNT_FILE = os.getenv('SERVICE_ACCOUNT_FILE', 'service_account_key.json')
    GDRIVE_SOURCE_FOLDER_ID = os.getenv('GDRIVE_SOURCE_FOLDER_ID', '')

    # Windows folder settings
    DESTINATION_FOLDER_PATH = os.getenv('DESTINATION_FOLDER_PATH', r'C:\Downloads\FromGDrive')

    # Transfer mode: 'copy' or 'move' (move will delete from Google Drive after download)
    TRANSFER_MODE = os.getenv('TRANSFER_MODE', 'copy')

    # File filter (optional): comma-separated file extensions to filter (e.g., '.pdf,.xlsx')
    FILE_EXTENSIONS_FILTER = os.getenv('FILE_EXTENSIONS_FILTER', '')

    # Service Account Email (for reference)
    SERVICE_ACCOUNT_EMAIL = 'drive-dasy-access@prj-nt-th-hub.iam.gserviceaccount.com'


class GoogleDriveClient:
    """Google Drive API client using service account"""

    SCOPES = [
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/drive.file'
    ]

    def __init__(self, config: Config):
        self.config = config
        self.service = None

    def authenticate(self) -> None:
        """Authenticate with Google Drive using service account"""
        try:
            credentials = service_account.Credentials.from_service_account_file(
                self.config.SERVICE_ACCOUNT_FILE,
                scopes=self.SCOPES
            )
            self.service = build('drive', 'v3', credentials=credentials)
            logger.info("Successfully authenticated with Google Drive")
        except Exception as e:
            logger.error(f"Failed to authenticate with Google Drive: {e}")
            raise

    def list_files(self, folder_id: str, file_extensions: List[str] = None) -> List[Dict[str, Any]]:
        """
        List all files in a Google Drive folder

        Args:
            folder_id: Google Drive folder ID
            file_extensions: Optional list of file extensions to filter (e.g., ['.pdf', '.xlsx'])

        Returns:
            List of file metadata dictionaries
        """
        files = []
        page_token = None

        try:
            query = f"'{folder_id}' in parents and trashed = false and mimeType != 'application/vnd.google-apps.folder'"

            while True:
                response = self.service.files().list(
                    q=query,
                    spaces='drive',
                    fields='nextPageToken, files(id, name, mimeType, size, createdTime, modifiedTime)',
                    pageToken=page_token,
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True
                ).execute()

                for file in response.get('files', []):
                    # Apply file extension filter if specified
                    if file_extensions:
                        file_ext = os.path.splitext(file['name'])[1].lower()
                        if file_ext not in [ext.lower() for ext in file_extensions]:
                            continue
                    files.append(file)

                page_token = response.get('nextPageToken')
                if not page_token:
                    break

            logger.info(f"Found {len(files)} file(s) in Google Drive folder")
            return files

        except Exception as e:
            logger.error(f"Failed to list files from Google Drive: {e}")
            raise

    def download_file(self, file_id: str, file_name: str, destination_path: str) -> Optional[str]:
        """
        Download a file from Google Drive

        Args:
            file_id: Google Drive file ID
            file_name: Name of the file
            destination_path: Local folder path to save the file

        Returns:
            Full path of the downloaded file, or None if failed
        """
        try:
            # Create destination directory if it doesn't exist
            os.makedirs(destination_path, exist_ok=True)

            file_path = os.path.join(destination_path, file_name)

            request = self.service.files().get_media(fileId=file_id, supportsAllDrives=True)
            file_handle = BytesIO()
            downloader = MediaIoBaseDownload(file_handle, request)

            done = False
            while not done:
                status, done = downloader.next_chunk()
                if status:
                    logger.debug(f"Download progress: {int(status.progress() * 100)}%")

            # Write to file
            with open(file_path, 'wb') as f:
                f.write(file_handle.getvalue())

            logger.info(f"Downloaded file: {file_name} -> {file_path}")
            return file_path

        except Exception as e:
            logger.error(f"Failed to download file {file_name}: {e}")
            return None

    def delete_file(self, file_id: str, file_name: str) -> bool:
        """
        Delete a file from Google Drive

        Args:
            file_id: Google Drive file ID
            file_name: Name of the file (for logging)

        Returns:
            True if deleted successfully, False otherwise
        """
        try:
            self.service.files().delete(fileId=file_id, supportsAllDrives=True).execute()
            logger.info(f"Deleted file from Google Drive: {file_name}")
            return True
        except Exception as e:
            logger.error(f"Failed to delete file {file_name} from Google Drive: {e}")
            return False


class FileTransferService:
    """Main service for transferring files from Google Drive to Windows folder"""

    def __init__(self, config: Config):
        self.config = config
        self.gdrive_client = GoogleDriveClient(config)

    def run(self) -> List[Dict[str, Any]]:
        """
        Execute the file transfer process

        Returns:
            List of dictionaries containing transfer results
        """
        results = []

        try:
            # Validate configuration
            if not self.config.GDRIVE_SOURCE_FOLDER_ID:
                logger.error("GDRIVE_SOURCE_FOLDER_ID is not configured")
                return results

            # Authenticate with Google Drive
            self.gdrive_client.authenticate()

            # Parse file extensions filter
            file_extensions = None
            if self.config.FILE_EXTENSIONS_FILTER:
                file_extensions = [
                    ext.strip() if ext.strip().startswith('.') else f'.{ext.strip()}'
                    for ext in self.config.FILE_EXTENSIONS_FILTER.split(',')
                    if ext.strip()
                ]
                logger.info(f"Filtering files by extensions: {file_extensions}")

            # List files in Google Drive folder
            files = self.gdrive_client.list_files(
                self.config.GDRIVE_SOURCE_FOLDER_ID,
                file_extensions
            )

            if not files:
                logger.warning("No files found to download")
                return results

            # Process each file
            for file_info in files:
                file_id = file_info['id']
                file_name = file_info['name']

                logger.info(f"Processing file: {file_name}")

                # Download file
                downloaded_path = self.gdrive_client.download_file(
                    file_id,
                    file_name,
                    self.config.DESTINATION_FOLDER_PATH
                )

                result = {
                    'file_name': file_name,
                    'file_id': file_id,
                    'downloaded_path': downloaded_path,
                    'deleted_from_gdrive': False
                }

                if downloaded_path:
                    # Delete from Google Drive if move mode
                    if self.config.TRANSFER_MODE.lower() == 'move':
                        result['deleted_from_gdrive'] = self.gdrive_client.delete_file(
                            file_id,
                            file_name
                        )

                results.append(result)

            successful_downloads = sum(1 for r in results if r['downloaded_path'])
            logger.info(f"Transfer complete. {successful_downloads}/{len(results)} file(s) downloaded.")

            return results

        except Exception as e:
            logger.error(f"Transfer process failed: {e}")
            raise


def main():
    """Main entry point"""
    config = Config()
    service = FileTransferService(config)

    try:
        results = service.run()

        if results:
            print("\n=== Download Summary ===")
            print(f"Transfer Mode: {config.TRANSFER_MODE.upper()}")
            print(f"Destination: {config.DESTINATION_FOLDER_PATH}")
            print("-" * 50)

            successful = 0
            for result in results:
                status = "OK" if result['downloaded_path'] else "FAILED"
                move_status = ""
                if config.TRANSFER_MODE.lower() == 'move':
                    move_status = " (deleted from GDrive)" if result['deleted_from_gdrive'] else " (NOT deleted)"

                print(f"  [{status}] {result['file_name']}{move_status}")
                if result['downloaded_path']:
                    successful += 1

            print("-" * 50)
            print(f"Total: {successful}/{len(results)} file(s) downloaded successfully")
        else:
            print("No files were downloaded")

    except Exception as e:
        logger.error(f"Process failed: {e}")
        raise


if __name__ == "__main__":
    main()
