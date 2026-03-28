"""
ITSEQ - SQL Server to AppSheet Processor

This script:
1. Queries SQL Server to find approved ITSEQ documents
2. Constructs document reference URL from SINSEI_CODE
3. Adds data to Google AppSheet Database (mine_taskx table)
4. Updates SQL Server to mark document as processed
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
    """Main configuration with environment support (ITSEQ-specific)"""

    # Current environment (ITSEQ uses separate environment setting from PRLOC)
    environment: Environment = field(default_factory=lambda: Environment(
        os.getenv('ITSEQ_ENVIRONMENT', 'test').lower()
    ))

    # Database configurations (ITSEQ-specific)
    db_configs: Dict[Environment, DatabaseConfig] = field(default_factory=lambda: {
        Environment.TEST: DatabaseConfig(
            server=os.getenv('ITSEQ_TEST_SQL_SERVER', 'test_server'),
            database=os.getenv('ITSEQ_TEST_SQL_DATABASE', 'DASY-FLII-TEST'),
            username=os.getenv('ITSEQ_TEST_SQL_USERNAME', 'test_user'),
            password=os.getenv('ITSEQ_TEST_SQL_PASSWORD', 'test_password'),
        ),
        Environment.PRODUCTION: DatabaseConfig(
            server=os.getenv('ITSEQ_PROD_SQL_SERVER', 'prod_server'),
            database=os.getenv('ITSEQ_PROD_SQL_DATABASE', 'DASY-FLII'),
            username=os.getenv('ITSEQ_PROD_SQL_USERNAME', 'prod_user'),
            password=os.getenv('ITSEQ_PROD_SQL_PASSWORD', 'prod_password'),
        ),
    })

    # AppSheet configuration for mine_taskx table (ITSEQ-specific)
    appsheet: AppSheetConfig = field(default_factory=lambda: AppSheetConfig(
        app_id=os.getenv('ITSEQ_APPSHEET_APP_ID', 'your_app_id'),
        api_key=os.getenv('ITSEQ_APPSHEET_API_KEY', 'your_api_key'),
        table_name=os.getenv('ITSEQ_APPSHEET_TABLE_NAME', 'mine_taskx'),
    ))

    # Document search settings
    auto_no_chr: str = field(default_factory=lambda: os.getenv('ITSEQ_AUTO_NO_CHR', 'ITSEQ'))
    auto_no: str = field(default_factory=lambda: os.getenv('ITSEQ_AUTO_NO', ''))  # For recovery mode

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
class ITSEQDocument:
    """ITSEQ document data"""
    sinsei_code: str            # SINSEI_CODE -> AppSheet: document_id
    auto_no: str = ""           # AUTO_NO -> AppSheet: application_no
    hakkou_syain_no: str = ""   # HAKKOU_SYAIN_NO -> AppSheet: applicant_code
    hakkou_syain_simei: str = ""  # HAKKOU_SYAIN_SIMEI -> AppSheet: applicant
    sinsei_syozoku_name: str = ""  # SINSEI_SYOZOKU_NAME -> AppSheet: placement
    syuyou_1: str = ""          # SYUYOU_1 -> AppSheet: subject
    syuyou_2: str = ""          # SYUYOU_2 -> AppSheet: approval_route_code
    syuyou_3: str = ""          # SYUYOU_3 -> AppSheet: company
    auto_no_chr: str = ""       # AUTO_NO_CHR -> AppSheet: doc_code
    refer_url: str = ""         # Constructed URL -> AppSheet: url_05_refer

    # URL template for document reference
    REFER_URL_TEMPLATE = (
        "http://dasy.niterraibcasia.com/DASY-FLII/ope/common/pdfprint.page"
        "?pdfPrintKbn=0&POPUPKBN=1&SHD_SCD={sinsei_code}&SHD_SRI_CODE=ITSEQ&SHD_SRI_HAN=1"
    )

    def build_refer_url(self) -> str:
        """Construct document reference URL from SINSEI_CODE"""
        return self.REFER_URL_TEMPLATE.format(sinsei_code=self.sinsei_code)

    def to_appsheet_row(self) -> Dict[str, Any]:
        """Convert to AppSheet row format for mine_taskx table"""
        row = {
            "document_id": self.sinsei_code,
            "application_no": self.auto_no,
            "applicant_code": self.hakkou_syain_no,
            "applicant": self.hakkou_syain_simei,
            "placement": self.sinsei_syozoku_name,
            "subject": self.syuyou_1,
            "approval_route_code": self.syuyou_2,
            "company": self.syuyou_3,
            "doc_code": self.auto_no_chr,
            "url_05_refer": self.refer_url,
        }
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

    def get_approved_documents(self, limit: int = 1) -> List[ITSEQDocument]:
        """
        Query to get approved ITSEQ documents

        Args:
            limit: Maximum number of documents to retrieve

        Returns:
            List of ITSEQDocument objects
        """
        select_columns = """
                SINSEI_CODE,
                AUTO_NO,
                HAKKOU_SYAIN_NO,
                HAKKOU_SYAIN_SIMEI,
                SINSEI_SYOZOKU_NAME,
                SYUYOU_1,
                SYUYOU_2,
                SYUYOU_3,
                AUTO_NO_CHR"""

        if self.config.auto_no:
            # Recovery mode: Search by exact document number
            query = f"""
                SELECT TOP {limit} {select_columns}
                FROM FLIISA.TR_SINSEI_DATA_HEADER
                WHERE AUTO_NO = ?
            """
            params = (self.config.auto_no,)
            logger.info(f"Recovery mode: Searching for AUTO_NO = '{self.config.auto_no}'")
        else:
            # Normal mode: Search for approved documents
            query = f"""
                SELECT TOP {limit} {select_columns}
                FROM FLIISA.TR_SINSEI_DATA_HEADER
                WHERE AUTO_NO_CHR = ?
                    AND JOUTAI_KBN = 1
                    AND (DATEPART(SECOND, INS_DATE) != 0 OR DATEPART(MILLISECOND, INS_DATE) != 0)
                    AND (DATEPART(SECOND, UPD_DATE) != 0 OR DATEPART(MILLISECOND, UPD_DATE) != 0)
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
                logger.info(f"Raw row data: SINSEI_CODE={row[0]}, AUTO_NO={row[1]}")

                doc = ITSEQDocument(
                    sinsei_code=row[0],
                    auto_no=row[1] or "",
                    hakkou_syain_no=row[2] or "",
                    hakkou_syain_simei=row[3] or "",
                    sinsei_syozoku_name=row[4] or "",
                    syuyou_1=row[5] or "",
                    syuyou_2=row[6] or "",
                    syuyou_3=row[7] or "",
                    auto_no_chr=row[8] or "",
                )
                doc.refer_url = doc.build_refer_url()
                documents.append(doc)

            logger.info(f"Found {len(documents)} approved document(s)")
            return documents

        except pyodbc.Error as e:
            logger.error(f"Error executing query: {e}")
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

            logger.info(f"Response status: {response.status_code}")
            logger.info(f"Response text: {response.text[:500] if response.text else '(empty)'}")

            response.raise_for_status()

            # AppSheet API returns HTTP 200 with empty body on success
            if not response.text or response.text.strip() == "":
                logger.info("AppSheet API returned 200 with empty body (success)")
                return {"success": True, "response": None}

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

class ITSEQService:
    """Main service for processing ITSEQ documents"""

    def __init__(self, config: Config):
        self.config = config
        self.sql_client = SQLServerClient(config)
        self.appsheet_client = AppSheetClient(config)

    def process_document(self, doc: ITSEQDocument) -> Dict[str, Any]:
        """
        Process a single document

        Args:
            doc: ITSEQDocument to process

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
            # Step 1: Add to AppSheet
            appsheet_result = self.appsheet_client.add_row(doc.to_appsheet_row())
            if appsheet_result["success"]:
                result["steps"].append({"step": "add_to_appsheet", "success": True})

                # Step 2: Update SQL Server (only if AppSheet was successful)
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
        logger.info("Starting ITSEQ Processor")
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
    service = ITSEQService(config)

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
