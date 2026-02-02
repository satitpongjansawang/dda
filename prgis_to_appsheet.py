"""
PRGIS - SQL Server to AppSheet Processor (Approval Chain)

This script:
1. Queries SQL Server to find approved PRGIS documents
2. Gets approval chain (up to 10 approvers) from TR_SINSEI_SYONIN
3. Gets document link from TR_SINSEI_DATA_LNK
4. Adds data to Google AppSheet Database (signed_box table)
5. Updates SQL Server to mark document as processed
"""

import os
import logging
from typing import Optional, List, Dict, Any
from dataclasses import dataclass, field
from enum import Enum

from dotenv import load_dotenv
import pyodbc
import requests

# Load environment variables from .env file
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# ===================== ENVIRONMENT CONFIGURATION =====================

class Environment(Enum):
    """Environment types"""
    TEST = "test"
    PRODUCTION = "production"


@dataclass
class DatabaseConfig:
    """Database configuration for each environment"""
    server: str
    database: str
    username: str
    password: str
    driver: str = "{ODBC Driver 17 for SQL Server}"


@dataclass
class AppSheetConfig:
    """AppSheet API configuration"""
    app_id: str
    api_key: str
    table_name: str
    base_url: str = "https://api.appsheet.com/api/v2/apps"


@dataclass
class Config:
    """Main configuration with environment support (PRGIS-specific)"""

    # Current environment (PRGIS uses separate environment setting from PRLOC)
    environment: Environment = field(default_factory=lambda: Environment(
        os.getenv('PRGIS_ENVIRONMENT', 'test').lower()
    ))

    # Database configurations (PRGIS-specific)
    db_configs: Dict[Environment, DatabaseConfig] = field(default_factory=lambda: {
        Environment.TEST: DatabaseConfig(
            server=os.getenv('PRGIS_TEST_SQL_SERVER', 'test_server'),
            database=os.getenv('PRGIS_TEST_SQL_DATABASE', 'DASY-FLII-TEST'),
            username=os.getenv('PRGIS_TEST_SQL_USERNAME', 'test_user'),
            password=os.getenv('PRGIS_TEST_SQL_PASSWORD', 'test_password'),
        ),
        Environment.PRODUCTION: DatabaseConfig(
            server=os.getenv('PRGIS_PROD_SQL_SERVER', 'prod_server'),
            database=os.getenv('PRGIS_PROD_SQL_DATABASE', 'DASY-FLII'),
            username=os.getenv('PRGIS_PROD_SQL_USERNAME', 'prod_user'),
            password=os.getenv('PRGIS_PROD_SQL_PASSWORD', 'prod_password'),
        ),
    })

    # AppSheet configuration for signed_box table (PRGIS-specific)
    appsheet: AppSheetConfig = field(default_factory=lambda: AppSheetConfig(
        app_id=os.getenv('PRGIS_APPSHEET_APP_ID', 'your_app_id'),
        api_key=os.getenv('PRGIS_APPSHEET_API_KEY', 'your_api_key'),
        table_name=os.getenv('PRGIS_APPSHEET_TABLE_NAME', 'signed_box'),
    ))

    # Document search settings
    auto_no_chr: str = field(default_factory=lambda: os.getenv('PRGIS_AUTO_NO_CHR', 'PRGIS'))
    auto_no: str = field(default_factory=lambda: os.getenv('PRGIS_AUTO_NO', ''))  # For recovery mode

    @property
    def db(self) -> DatabaseConfig:
        """Get current database configuration based on environment"""
        return self.db_configs[self.environment]

    def __post_init__(self):
        logger.info(f"Configuration loaded for environment: {self.environment.value}")
        logger.info(f"Database: {self.db.database} on {self.db.server}")
        logger.info(f"AppSheet table: {self.appsheet.table_name}")


# ===================== DATA MODELS =====================

@dataclass
class ApprovalInfo:
    """Approval information for a single approver"""
    approver_name: str = ""
    approval_date: str = ""


@dataclass
class PRGISDocument:
    """PRGIS document data with approval chain"""
    sinsei_code: str
    auto_no: str = ""           # AUTO_NO → AppSheet: dasy_auto_no
    sheet_url: str = ""         # LNK1 value → AppSheet: sheet_url
    approvers: List[ApprovalInfo] = field(default_factory=list)  # Up to 10 approvers

    def to_appsheet_row(self) -> Dict[str, Any]:
        """Convert to AppSheet row format for signed_box table"""
        row = {
            "dasy_auto_no": self.auto_no,
            "sheet_url": self.sheet_url,
            "is_force_stamp": False,
            "is_active": True,
        }

        # Add approvers (up to 10)
        for i in range(10):
            approver_key = f"approver{i + 1}"
            time_key = f"approver_time{i + 1}"

            if i < len(self.approvers):
                row[approver_key] = self.approvers[i].approver_name
                row[time_key] = self.approvers[i].approval_date
            else:
                row[approver_key] = ""
                row[time_key] = ""

        return row


# ===================== SQL SERVER CLIENT =====================

class SQLServerClient:
    """SQL Server database client"""

    def __init__(self, config: Config):
        self.config = config
        self.connection: Optional[pyodbc.Connection] = None

    def connect(self) -> None:
        """Establish connection to SQL Server"""
        db = self.config.db
        connection_string = (
            f"DRIVER={db.driver};"
            f"SERVER={db.server};"
            f"DATABASE={db.database};"
            f"UID={db.username};"
            f"PWD={db.password};"
            "TrustServerCertificate=yes;"
        )
        try:
            self.connection = pyodbc.connect(connection_string)
            logger.info(f"Connected to SQL Server: {db.database}")
        except pyodbc.Error as e:
            logger.error(f"Failed to connect to SQL Server: {e}")
            raise

    def disconnect(self) -> None:
        """Close the database connection"""
        if self.connection:
            self.connection.close()
            logger.info("Disconnected from SQL Server")

    def get_approved_documents(self, limit: int = 1) -> List[PRGISDocument]:
        """
        Query to get approved PRGIS documents

        Args:
            limit: Maximum number of documents to retrieve

        Returns:
            List of PRGISDocument objects
        """
        if self.config.auto_no:
            # Recovery mode: Search by exact document number
            query = f"""
                SELECT TOP {limit} SINSEI_CODE, AUTO_NO
                FROM FLIISA.TR_SINSEI_DATA_HEADER
                WHERE AUTO_NO = ?
            """
            params = (self.config.auto_no,)
            logger.info(f"Recovery mode: Searching for AUTO_NO = '{self.config.auto_no}'")
        else:
            # Normal mode: Search for approved documents
            query = f"""
                SELECT TOP {limit} SINSEI_CODE, AUTO_NO
                FROM FLIISA.TR_SINSEI_DATA_HEADER
                WHERE AUTO_NO_CHR = ?
                    AND JOUTAI_KBN = 1
                    AND DATEPART(MILLISECOND, INS_DATE) != 0
                    AND DATEPART(MILLISECOND, UPD_DATE) != 0
                    AND DATEPART(SECOND, INS_DATE) != 0
                    AND DATEPART(SECOND, UPD_DATE) != 0
                ORDER BY UPD_DATE ASC
            """
            params = (self.config.auto_no_chr,)
            logger.info(f"Normal mode: Searching for AUTO_NO_CHR = '{self.config.auto_no_chr}'")

        try:
            cursor = self.connection.cursor()
            cursor.execute(query, params)
            rows = cursor.fetchall()

            documents = []
            for row in rows:
                doc = PRGISDocument(
                    sinsei_code=row[0],
                    auto_no=row[1] or "",
                )
                documents.append(doc)

            logger.info(f"Found {len(documents)} approved document(s)")
            return documents

        except pyodbc.Error as e:
            logger.error(f"Error executing query: {e}")
            raise

    def get_document_link(self, sinsei_code: str) -> Optional[str]:
        """
        Get LNK1 value from TR_SINSEI_DATA_LNK

        Args:
            sinsei_code: Document SINSEI_CODE

        Returns:
            LNK1 value (sheet URL) or None
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

            if row and row[0]:
                logger.info(f"Found LNK1 for {sinsei_code}")
                return row[0]
            else:
                logger.warning(f"No LNK1 found for {sinsei_code}")
                return None

        except pyodbc.Error as e:
            logger.error(f"Error getting document link: {e}")
            raise

    def get_approval_chain(self, sinsei_code: str) -> List[ApprovalInfo]:
        """
        Get approval chain from TR_SINSEI_SYONIN

        Args:
            sinsei_code: Document SINSEI_CODE

        Returns:
            List of ApprovalInfo (up to 10 approvers)
        """
        query = """
            SELECT
                SYONIN_SIMEI,
                SYONIN_DATE
            FROM FLIISA.TR_SINSEI_SYONIN
            WHERE SINSEI_CODE = ?
            ORDER BY SYONIN_STEP_NO, SYONIN_NO
        """

        try:
            cursor = self.connection.cursor()
            cursor.execute(query, (sinsei_code,))
            rows = cursor.fetchall()

            approvers = []
            for row in rows[:10]:  # Limit to 10 approvers
                approval_date = ""
                if row[1]:
                    # Format date as string
                    approval_date = str(row[1])

                approvers.append(ApprovalInfo(
                    approver_name=row[0] or "",
                    approval_date=approval_date,
                ))

            logger.info(f"Found {len(approvers)} approver(s) for {sinsei_code}")
            return approvers

        except pyodbc.Error as e:
            logger.error(f"Error getting approval chain: {e}")
            raise

    def update_document_status(self, sinsei_code: str) -> bool:
        """
        Update document to mark as processed
        (Set milliseconds and seconds to 0 to indicate processed)

        Args:
            sinsei_code: Document SINSEI_CODE

        Returns:
            True if update successful
        """
        query = """
            UPDATE FLIISA.TR_SINSEI_DATA_HEADER
            SET INS_DATE = DATEADD(MILLISECOND, -DATEPART(MILLISECOND, INS_DATE),
                           DATEADD(SECOND, -DATEPART(SECOND, INS_DATE), INS_DATE)),
                UPD_DATE = DATEADD(MILLISECOND, -DATEPART(MILLISECOND, UPD_DATE),
                           DATEADD(SECOND, -DATEPART(SECOND, UPD_DATE), UPD_DATE))
            WHERE SINSEI_CODE = ?
        """

        try:
            cursor = self.connection.cursor()
            cursor.execute(query, (sinsei_code,))
            self.connection.commit()

            if cursor.rowcount > 0:
                logger.info(f"Updated document status for {sinsei_code}")
                return True
            else:
                logger.warning(f"No rows updated for {sinsei_code}")
                return False

        except pyodbc.Error as e:
            logger.error(f"Error updating document status: {e}")
            self.connection.rollback()
            raise


# ===================== APPSHEET CLIENT =====================

class AppSheetClient:
    """Google AppSheet API client"""

    def __init__(self, config: Config):
        self.config = config
        self.app_config = config.appsheet

    @property
    def api_url(self) -> str:
        """Get the API endpoint URL"""
        return f"{self.app_config.base_url}/{self.app_config.app_id}/tables/{self.app_config.table_name}/Action"

    @property
    def headers(self) -> Dict[str, str]:
        """Get API request headers"""
        return {
            "ApplicationAccessKey": self.app_config.api_key,
            "Content-Type": "application/json"
        }

    def add_row(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Add a new row to AppSheet table

        Args:
            data: Row data as dictionary

        Returns:
            API response
        """
        payload = {
            "Action": "Add",
            "Properties": {
                "Locale": "en-US",
                "Timezone": "SE Asia Standard Time"
            },
            "Rows": [data]
        }

        try:
            logger.info(f"Adding row to AppSheet: {self.app_config.table_name}")
            logger.info(f"API URL: {self.api_url}")
            logger.info(f"Payload: {payload}")

            response = requests.post(
                self.api_url,
                json=payload,
                headers=self.headers,
                timeout=30
            )

            # Log response details for debugging
            logger.info(f"Response status: {response.status_code}")
            logger.info(f"Response text: {response.text[:500] if response.text else '(empty)'}")

            response.raise_for_status()

            # Check if response is empty
            if not response.text:
                logger.warning("Empty response from AppSheet API")
                return {"success": False, "error": "Empty response from API"}

            result = response.json()
            logger.info("Successfully added row to AppSheet")
            return {"success": True, "response": result}

        except requests.exceptions.RequestException as e:
            logger.error(f"Error adding row to AppSheet: {e}")
            return {"success": False, "error": str(e)}
        except ValueError as e:
            logger.error(f"Error parsing JSON response: {e}")
            return {"success": False, "error": f"Invalid JSON response: {e}"}

    def test_connection(self) -> bool:
        """
        Test AppSheet API connection

        Returns:
            True if connection successful
        """
        payload = {
            "Action": "Find",
            "Properties": {
                "Locale": "en-US"
            },
            "Rows": []
        }

        try:
            response = requests.post(
                self.api_url,
                json=payload,
                headers=self.headers,
                timeout=10
            )
            response.raise_for_status()
            logger.info("AppSheet API connection successful")
            return True
        except requests.exceptions.RequestException as e:
            logger.error(f"AppSheet API connection failed: {e}")
            return False


# ===================== MAIN SERVICE =====================

class PRGISService:
    """Main service for processing PRGIS documents"""

    def __init__(self, config: Config):
        self.config = config
        self.sql_client = SQLServerClient(config)
        self.appsheet_client = AppSheetClient(config)

    def process_document(self, doc: PRGISDocument) -> Dict[str, Any]:
        """
        Process a single document

        Args:
            doc: PRGISDocument to process

        Returns:
            Processing result
        """
        result = {
            "sinsei_code": doc.sinsei_code,
            "auto_no": doc.auto_no,
            "success": False,
            "steps": []
        }

        try:
            # Step 1: Get document link (LNK1)
            doc.sheet_url = self.sql_client.get_document_link(doc.sinsei_code) or ""
            if doc.sheet_url:
                result["steps"].append({"step": "get_link", "success": True})
            else:
                result["steps"].append({"step": "get_link", "success": False, "error": "No LNK1 found"})

            # Step 2: Get approval chain
            doc.approvers = self.sql_client.get_approval_chain(doc.sinsei_code)
            result["steps"].append({
                "step": "get_approvers",
                "success": True,
                "count": len(doc.approvers)
            })

            # Step 3: Add to AppSheet
            appsheet_result = self.appsheet_client.add_row(doc.to_appsheet_row())
            if appsheet_result["success"]:
                result["steps"].append({"step": "add_to_appsheet", "success": True})

                # Step 4: Update SQL Server (only if AppSheet was successful)
                update_success = self.sql_client.update_document_status(doc.sinsei_code)
                result["steps"].append({"step": "update_sql", "success": update_success})

                if update_success:
                    result["success"] = True
            else:
                result["steps"].append({
                    "step": "add_to_appsheet",
                    "success": False,
                    "error": appsheet_result.get("error")
                })

            return result

        except Exception as e:
            logger.error(f"Error processing document {doc.sinsei_code}: {e}")
            result["error"] = str(e)
            return result

    def run(self, limit: int = 1) -> List[Dict[str, Any]]:
        """
        Execute the main processing workflow

        Args:
            limit: Maximum number of documents to process

        Returns:
            List of processing results
        """
        results = []

        logger.info("=" * 60)
        logger.info("Starting PRGIS Processor (Approval Chain)")
        logger.info(f"Environment: {self.config.environment.value}")
        logger.info(f"Document type: {self.config.auto_no_chr}")
        logger.info("=" * 60)

        try:
            # Connect to SQL Server
            self.sql_client.connect()

            # Test AppSheet connection
            if not self.appsheet_client.test_connection():
                logger.error("AppSheet API connection failed. Aborting.")
                return results

            # Get approved documents
            documents = self.sql_client.get_approved_documents(limit=limit)

            if not documents:
                logger.info("No documents to process")
                return results

            # Process each document
            for doc in documents:
                logger.info(f"Processing document: {doc.auto_no} ({doc.sinsei_code})")
                result = self.process_document(doc)
                results.append(result)

                if result["success"]:
                    logger.info(f"Document {doc.auto_no} processed successfully")
                else:
                    logger.warning(f"Document {doc.auto_no} processing failed")

            logger.info("=" * 60)
            logger.info(f"Processing complete. {sum(1 for r in results if r['success'])}/{len(results)} successful")
            logger.info("=" * 60)

            return results

        finally:
            self.sql_client.disconnect()


# ===================== ENTRY POINT =====================

def main():
    """Main entry point"""
    config = Config()
    service = PRGISService(config)

    try:
        results = service.run(limit=1)

        print("\n" + "=" * 60)
        print("Processing Results:")
        print("=" * 60)

        for result in results:
            status = "SUCCESS" if result["success"] else "FAILED"
            print(f"\n  AUTO_NO: {result['auto_no']} - {status}")
            print(f"  SINSEI_CODE: {result['sinsei_code']}")
            for step in result.get("steps", []):
                step_status = "OK" if step["success"] else "FAIL"
                extra = ""
                if "count" in step:
                    extra = f" ({step['count']} items)"
                if "error" in step:
                    extra = f" - {step['error']}"
                print(f"    - {step['step']}: {step_status}{extra}")

        successful = sum(1 for r in results if r["success"])
        print(f"\nTotal: {successful}/{len(results)} document(s) processed successfully")

    except Exception as e:
        logger.error(f"Process failed: {e}")
        raise


if __name__ == "__main__":
    main()
