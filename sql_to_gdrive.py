"""
SQL Server to Google Drive File Transfer Script

This script:
1. Queries SQL Server to find records matching specific conditions
2. Gets SINSEI_CODE and INTERNAL_NAME for file identification
3. Copies the identified files to Google Drive
"""

import os
import logging
from typing import Optional, List, Tuple

from dotenv import load_dotenv
import pyodbc
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

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
    # AUTO_NO_CHR: Document prefix (e.g., 'ITSEQ') - searches documents starting with this prefix
    # AUTO_NO: Specific document number (e.g., 'ITSEQ-2601-00002') - for recovery mode
    AUTO_NO_CHR = os.getenv('AUTO_NO_CHR', 'ITSEQ')
    AUTO_NO = os.getenv('AUTO_NO', '')  # Empty means use AUTO_NO_CHR mode

    # File paths
    BASE_FILE_PATH = os.getenv('BASE_FILE_PATH', r'C:\Myapp\DASY-FLII\WEB-INF\AppendFiles')

    # Google Drive settings
    SERVICE_ACCOUNT_FILE = os.getenv('SERVICE_ACCOUNT_FILE', 'service_account_key.json')
    GOOGLE_DRIVE_FOLDER_ID = os.getenv('GOOGLE_DRIVE_FOLDER_ID', '1t-iToHHsud8uZVer_W9K7slmel-clySw')

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

    def get_sinsei_code(self, auto_no: str = '', auto_no_chr: str = 'ITSEQ') -> Optional[str]:
        """
        Query to get SINSEI_CODE from TR_SINSEI_DATA_HEADER

        Args:
            auto_no: Specific document number (e.g., 'ITSEQ-2601-00002') for recovery mode
            auto_no_chr: Document prefix (e.g., 'ITSEQ') for normal mode

        Returns the first matching record based on conditions.
        If auto_no is provided, searches by exact document number (recovery mode).
        Otherwise, searches by prefix (normal mode).
        """
        if auto_no:
            # Recovery mode: Search by exact document number
            query = """
                SELECT TOP 1 SINSEI_CODE
                FROM FLIISA.TR_SINSEI_DATA_HEADER
                WHERE AUTO_NO = ?
            """
            params = (auto_no,)
            logger.info(f"Recovery mode: Searching for document AUTO_NO = '{auto_no}'")
        else:
            # Normal mode: Search by prefix with conditions
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
            params = (auto_no_chr,)
            logger.info(f"Normal mode: Searching for documents with prefix AUTO_NO_CHR = '{auto_no_chr}'")

        try:
            cursor = self.connection.cursor()
            cursor.execute(query, params)
            row = cursor.fetchone()
            if row:
                sinsei_code = row[0]
                logger.info(f"Found SINSEI_CODE: {sinsei_code}")
                return sinsei_code
            else:
                logger.warning("No matching SINSEI_CODE found")
                return None
        except pyodbc.Error as e:
            logger.error(f"Error executing query: {e}")
            raise

    def get_internal_names(self, sinsei_code: str) -> List[str]:
        """
        Query to get INTERNAL_NAME(s) from TR_SINSEI_DATA_TMT
        using the given SINSEI_CODE
        """
        query = """
            SELECT INTERNAL_NAME
            FROM FLIISA.TR_SINSEI_DATA_TMT
            WHERE SINSEI_CODE = ?
        """
        try:
            cursor = self.connection.cursor()
            cursor.execute(query, (sinsei_code,))
            rows = cursor.fetchall()
            internal_names = [row[0] for row in rows if row[0]]
            logger.info(f"Found {len(internal_names)} INTERNAL_NAME(s) for SINSEI_CODE: {sinsei_code}")
            return internal_names
        except pyodbc.Error as e:
            logger.error(f"Error executing query: {e}")
            raise


class GoogleDriveClient:
    """Google Drive API client using service account"""

    SCOPES = ['https://www.googleapis.com/auth/drive.file']

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

    def upload_file(self, file_path: str, folder_id: str) -> Optional[str]:
        """
        Upload a file to Google Drive

        Args:
            file_path: Local path to the file
            folder_id: Google Drive folder ID to upload to

        Returns:
            File ID of the uploaded file, or None if failed
        """
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            return None

        file_name = os.path.basename(file_path)
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }

        try:
            media = MediaFileUpload(file_path, resumable=True)
            file = self.service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, name',
                supportsAllDrives=True
            ).execute()

            logger.info(f"Uploaded file: {file_name} (ID: {file.get('id')})")
            return file.get('id')
        except Exception as e:
            logger.error(f"Failed to upload file {file_name}: {e}")
            return None


class FileTransferService:
    """Main service for transferring files from SQL Server records to Google Drive"""

    def __init__(self, config: Config):
        self.config = config
        self.sql_client = SQLServerClient(config)
        self.gdrive_client = GoogleDriveClient(config)

    def run(self) -> List[Tuple[str, str]]:
        """
        Execute the file transfer process

        Returns:
            List of tuples containing (file_name, google_drive_file_id)
        """
        uploaded_files = []

        try:
            # Connect to SQL Server
            self.sql_client.connect()

            # Get SINSEI_CODE (uses AUTO_NO for recovery mode, or AUTO_NO_CHR for normal mode)
            sinsei_code = self.sql_client.get_sinsei_code(
                auto_no=self.config.AUTO_NO,
                auto_no_chr=self.config.AUTO_NO_CHR
            )
            if not sinsei_code:
                logger.warning("No records found to process")
                return uploaded_files

            # Get INTERNAL_NAME(s)
            internal_names = self.sql_client.get_internal_names(sinsei_code)
            if not internal_names:
                logger.warning(f"No files found for SINSEI_CODE: {sinsei_code}")
                return uploaded_files

            # Authenticate with Google Drive
            self.gdrive_client.authenticate()

            # Process each file
            for internal_name in internal_names:
                # Construct file path
                file_path = os.path.join(
                    self.config.BASE_FILE_PATH,
                    sinsei_code,
                    internal_name
                )

                logger.info(f"Processing file: {file_path}")

                # Upload to Google Drive
                file_id = self.gdrive_client.upload_file(
                    file_path,
                    self.config.GOOGLE_DRIVE_FOLDER_ID
                )

                if file_id:
                    uploaded_files.append((internal_name, file_id))

            logger.info(f"Transfer complete. {len(uploaded_files)} file(s) uploaded.")
            return uploaded_files

        finally:
            self.sql_client.disconnect()


def main():
    """Main entry point"""
    config = Config()
    service = FileTransferService(config)

    try:
        results = service.run()

        if results:
            print("\n=== Upload Summary ===")
            for file_name, file_id in results:
                print(f"  - {file_name}: {file_id}")
            print(f"\nTotal: {len(results)} file(s) uploaded to Google Drive")
        else:
            print("No files were uploaded")

    except Exception as e:
        logger.error(f"Process failed: {e}")
        raise


if __name__ == "__main__":
    main()
