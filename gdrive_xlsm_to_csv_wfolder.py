"""
Google Drive XLSM to CSV Windows Folder Transfer Script

This script:
1. Queries SQL Server to find approved documents (JOUTAI_KBN = 1)
2. Gets Google Drive link from TR_SINSEI_DATA_LNK table
3. Downloads the .xlsm file from Google Drive
4. Converts .xlsm to .csv format
5. Saves the .csv file to a Windows folder
"""

import os
import re
import logging
from typing import Optional, List, Dict, Any
from io import BytesIO

from dotenv import load_dotenv
import pyodbc
import pandas as pd
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
    # SQL Server connection settings
    SQL_SERVER = os.getenv('SQL_SERVER', 'your_server_name')
    SQL_DATABASE = os.getenv('SQL_DATABASE', 'DASY-FLII')
    SQL_USERNAME = os.getenv('SQL_USERNAME', 'your_username')
    SQL_PASSWORD = os.getenv('SQL_PASSWORD', 'your_password')
    SQL_DRIVER = os.getenv('SQL_DRIVER', '{ODBC Driver 17 for SQL Server}')

    # Document search settings
    AUTO_NO_CHR = os.getenv('AUTO_NO_CHR', 'ITSEQ')

    # Google Drive settings
    SERVICE_ACCOUNT_FILE = os.getenv('SERVICE_ACCOUNT_FILE', 'service_account_key.json')

    # Windows folder settings for CSV output
    CSV_DESTINATION_FOLDER = os.getenv('CSV_DESTINATION_FOLDER', r'C:\Downloads\CSV_Output')

    # Transfer mode: 'copy' or 'move' (move will delete from Google Drive after download)
    TRANSFER_MODE = os.getenv('XLSM_TRANSFER_MODE', 'copy')

    # Keep original xlsm file after conversion (true/false)
    KEEP_XLSM = os.getenv('KEEP_XLSM', 'false').lower() == 'true'

    # CSV encoding
    CSV_ENCODING = os.getenv('CSV_ENCODING', 'utf-8-sig')

    # Sheet name to convert (empty = first sheet)
    SHEET_NAME = os.getenv('XLSM_SHEET_NAME', '')

    # Service Account Email (for reference)
    SERVICE_ACCOUNT_EMAIL = 'drive-dasy-access@prj-nt-th-hub.iam.gserviceaccount.com'


class SQLServerClient:
    """SQL Server database client"""

    def __init__(self, config: Config):
        self.config = config
        self.connection: Optional[pyodbc.Connection] = None

    def connect(self) -> None:
        """Establish connection to SQL Server"""
        connection_string = (
            f"DRIVER={self.config.SQL_DRIVER};"
            f"SERVER={self.config.SQL_SERVER};"
            f"DATABASE={self.config.SQL_DATABASE};"
            f"UID={self.config.SQL_USERNAME};"
            f"PWD={self.config.SQL_PASSWORD};"
            "TrustServerCertificate=yes;"
        )
        try:
            self.connection = pyodbc.connect(connection_string)
            logger.info("Successfully connected to SQL Server")
        except pyodbc.Error as e:
            logger.error(f"Failed to connect to SQL Server: {e}")
            raise

    def disconnect(self) -> None:
        """Close the database connection"""
        if self.connection:
            self.connection.close()
            logger.info("Disconnected from SQL Server")

    def get_approved_sinsei_code(self, auto_no_chr: str = 'ITSEQ') -> Optional[str]:
        """
        Query to get SINSEI_CODE from approved documents in TR_SINSEI_DATA_HEADER

        Args:
            auto_no_chr: Document prefix (e.g., 'ITSEQ')

        Returns:
            SINSEI_CODE of the first approved document, or None if not found
        """
        query = """
            SELECT TOP 1 SINSEI_CODE
            FROM FLIISA.TR_SINSEI_DATA_HEADER
            WHERE AUTO_NO_CHR = ?
                AND JOUTAI_KBN = 1
                AND DATEPART(MILLISECOND, INS_DATE) != 0
                AND DATEPART(MILLISECOND, UPD_DATE) != 0
                AND DATEPART(SECOND, INS_DATE) != 0
                AND DATEPART(SECOND, UPD_DATE) != 0
            ORDER BY UPD_DATE ASC
        """

        try:
            cursor = self.connection.cursor()
            cursor.execute(query, (auto_no_chr,))
            row = cursor.fetchone()
            if row:
                sinsei_code = row[0]
                logger.info(f"Found approved SINSEI_CODE: {sinsei_code}")
                return sinsei_code
            else:
                logger.warning(f"No approved document found for AUTO_NO_CHR = '{auto_no_chr}'")
                return None
        except pyodbc.Error as e:
            logger.error(f"Error executing query: {e}")
            raise

    def get_gdrive_link(self, sinsei_code: str) -> Optional[str]:
        """
        Query to get Google Drive link from TR_SINSEI_DATA_LNK

        Args:
            sinsei_code: SINSEI_CODE to search for

        Returns:
            Google Drive link (KOUMOKU_VALUE) or None if not found
        """
        query = """
            SELECT KOUMOKU_VALUE
            FROM FLIISA.TR_SINSEI_DATA_LNK
            WHERE SINSEI_CODE = ?
                AND KOUMOKU_KEY = 'LNK1'
        """

        try:
            cursor = self.connection.cursor()
            cursor.execute(query, (sinsei_code,))
            row = cursor.fetchone()
            if row:
                gdrive_link = row[0]
                logger.info(f"Found Google Drive link: {gdrive_link}")
                return gdrive_link
            else:
                logger.warning(f"No Google Drive link found for SINSEI_CODE = '{sinsei_code}'")
                return None
        except pyodbc.Error as e:
            logger.error(f"Error executing query: {e}")
            raise


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

    @staticmethod
    def extract_file_id_from_url(url: str) -> Optional[str]:
        """
        Extract Google Drive file ID from various URL formats

        Supported formats:
        - https://drive.google.com/file/d/FILE_ID/view
        - https://drive.google.com/open?id=FILE_ID
        - https://docs.google.com/spreadsheets/d/FILE_ID/edit
        - Direct file ID

        Args:
            url: Google Drive URL or file ID

        Returns:
            File ID or None if not found
        """
        if not url:
            return None

        # Pattern for /d/FILE_ID/ format
        pattern1 = r'/d/([a-zA-Z0-9_-]+)'
        match = re.search(pattern1, url)
        if match:
            return match.group(1)

        # Pattern for ?id=FILE_ID format
        pattern2 = r'[?&]id=([a-zA-Z0-9_-]+)'
        match = re.search(pattern2, url)
        if match:
            return match.group(1)

        # If no pattern matches, assume it's already a file ID
        if re.match(r'^[a-zA-Z0-9_-]+$', url):
            return url

        return None

    def get_file_metadata(self, file_id: str) -> Optional[Dict[str, Any]]:
        """
        Get file metadata from Google Drive

        Args:
            file_id: Google Drive file ID

        Returns:
            File metadata dictionary or None if failed
        """
        try:
            file_metadata = self.service.files().get(
                fileId=file_id,
                fields='id, name, mimeType, size',
                supportsAllDrives=True
            ).execute()
            logger.info(f"File metadata: {file_metadata.get('name')} ({file_metadata.get('mimeType')})")
            return file_metadata
        except Exception as e:
            logger.error(f"Failed to get file metadata: {e}")
            return None

    def download_file(self, file_id: str, destination_path: str, file_name: str = None) -> Optional[str]:
        """
        Download a file from Google Drive

        Args:
            file_id: Google Drive file ID
            destination_path: Local folder path to save the file
            file_name: Optional file name (if None, uses original name)

        Returns:
            Full path of the downloaded file, or None if failed
        """
        try:
            # Get file metadata if name not provided
            if not file_name:
                metadata = self.get_file_metadata(file_id)
                if metadata:
                    file_name = metadata.get('name', f'{file_id}.xlsm')
                else:
                    file_name = f'{file_id}.xlsm'

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
            logger.error(f"Failed to download file: {e}")
            return None

    def delete_file(self, file_id: str) -> bool:
        """
        Delete a file from Google Drive

        Args:
            file_id: Google Drive file ID

        Returns:
            True if deleted successfully, False otherwise
        """
        try:
            self.service.files().delete(fileId=file_id, supportsAllDrives=True).execute()
            logger.info(f"Deleted file from Google Drive: {file_id}")
            return True
        except Exception as e:
            logger.error(f"Failed to delete file from Google Drive: {e}")
            return False


class ExcelConverter:
    """Excel file converter"""

    @staticmethod
    def xlsm_to_csv(
        xlsm_path: str,
        csv_path: str,
        sheet_name: str = None,
        encoding: str = 'utf-8-sig'
    ) -> bool:
        """
        Convert .xlsm file to .csv

        Args:
            xlsm_path: Path to the .xlsm file
            csv_path: Output path for the .csv file
            sheet_name: Name of the sheet to convert (None = first sheet)
            encoding: CSV encoding (default: utf-8-sig for Excel compatibility)

        Returns:
            True if conversion successful, False otherwise
        """
        try:
            # Read Excel file
            if sheet_name:
                df = pd.read_excel(xlsm_path, sheet_name=sheet_name, engine='openpyxl')
            else:
                df = pd.read_excel(xlsm_path, sheet_name=0, engine='openpyxl')

            # Create output directory if needed
            os.makedirs(os.path.dirname(csv_path), exist_ok=True)

            # Save as CSV
            df.to_csv(csv_path, index=False, encoding=encoding)
            logger.info(f"Converted to CSV: {csv_path}")
            return True

        except Exception as e:
            logger.error(f"Failed to convert XLSM to CSV: {e}")
            return False


class FileTransferService:
    """Main service for transferring and converting files"""

    def __init__(self, config: Config):
        self.config = config
        self.sql_client = SQLServerClient(config)
        self.gdrive_client = GoogleDriveClient(config)
        self.converter = ExcelConverter()

    def run(self) -> Dict[str, Any]:
        """
        Execute the file transfer and conversion process

        Returns:
            Dictionary containing process results
        """
        result = {
            'sinsei_code': None,
            'gdrive_link': None,
            'file_id': None,
            'xlsm_path': None,
            'csv_path': None,
            'success': False,
            'deleted_from_gdrive': False
        }

        try:
            # Connect to SQL Server
            self.sql_client.connect()

            # Step 1: Get approved SINSEI_CODE
            sinsei_code = self.sql_client.get_approved_sinsei_code(self.config.AUTO_NO_CHR)
            if not sinsei_code:
                logger.warning("No approved documents found to process")
                return result
            result['sinsei_code'] = sinsei_code

            # Step 2: Get Google Drive link
            gdrive_link = self.sql_client.get_gdrive_link(sinsei_code)
            if not gdrive_link:
                logger.warning(f"No Google Drive link found for SINSEI_CODE: {sinsei_code}")
                return result
            result['gdrive_link'] = gdrive_link

            # Step 3: Extract file ID from link
            file_id = self.gdrive_client.extract_file_id_from_url(gdrive_link)
            if not file_id:
                logger.error(f"Could not extract file ID from link: {gdrive_link}")
                return result
            result['file_id'] = file_id
            logger.info(f"Extracted file ID: {file_id}")

            # Authenticate with Google Drive
            self.gdrive_client.authenticate()

            # Step 4: Download .xlsm file
            xlsm_path = self.gdrive_client.download_file(
                file_id,
                self.config.CSV_DESTINATION_FOLDER
            )
            if not xlsm_path:
                logger.error("Failed to download file from Google Drive")
                return result
            result['xlsm_path'] = xlsm_path

            # Step 5: Convert .xlsm to .csv
            csv_filename = os.path.splitext(os.path.basename(xlsm_path))[0] + '.csv'
            csv_path = os.path.join(self.config.CSV_DESTINATION_FOLDER, csv_filename)

            sheet_name = self.config.SHEET_NAME if self.config.SHEET_NAME else None
            conversion_success = self.converter.xlsm_to_csv(
                xlsm_path,
                csv_path,
                sheet_name=sheet_name,
                encoding=self.config.CSV_ENCODING
            )

            if not conversion_success:
                logger.error("Failed to convert XLSM to CSV")
                return result
            result['csv_path'] = csv_path

            # Step 6: Clean up xlsm file if not keeping
            if not self.config.KEEP_XLSM and os.path.exists(xlsm_path):
                os.remove(xlsm_path)
                logger.info(f"Removed temporary XLSM file: {xlsm_path}")
                result['xlsm_path'] = None

            # Step 7: Delete from Google Drive if move mode
            if self.config.TRANSFER_MODE.lower() == 'move':
                result['deleted_from_gdrive'] = self.gdrive_client.delete_file(file_id)

            result['success'] = True
            logger.info("Process completed successfully")
            return result

        except Exception as e:
            logger.error(f"Process failed: {e}")
            raise

        finally:
            self.sql_client.disconnect()


def main():
    """Main entry point"""
    config = Config()
    service = FileTransferService(config)

    try:
        result = service.run()

        print("\n=== Process Summary ===")
        print(f"SINSEI_CODE: {result['sinsei_code'] or 'Not found'}")
        print(f"Google Drive Link: {result['gdrive_link'] or 'Not found'}")
        print(f"File ID: {result['file_id'] or 'Not extracted'}")
        print("-" * 50)

        if result['success']:
            print(f"[OK] CSV file created: {result['csv_path']}")
            if result['deleted_from_gdrive']:
                print(f"[OK] Original file deleted from Google Drive")
            if result['xlsm_path']:
                print(f"[INFO] XLSM file kept: {result['xlsm_path']}")
        else:
            print("[FAILED] Process did not complete successfully")

    except Exception as e:
        logger.error(f"Process failed: {e}")
        raise


if __name__ == "__main__":
    main()
