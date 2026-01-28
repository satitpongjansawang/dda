"""
PR Local Purchase - SQL Server to AppSheet Processor

This script:
1. Queries SQL Server to find approved PR Local Purchase documents
2. Gets document details (SINSEI_CODE, Google Drive link, etc.)
3. Adds data to Google AppSheet Database via API
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
    """Main configuration with environment support"""

    # Current environment
    environment: Environment = field(default_factory=lambda: Environment(
        os.getenv('ENVIRONMENT', 'test').lower()
    ))

    # Database configurations
    db_configs: Dict[Environment, DatabaseConfig] = field(default_factory=lambda: {
        Environment.TEST: DatabaseConfig(
            server=os.getenv('TEST_SQL_SERVER', 'test_server'),
            database=os.getenv('TEST_SQL_DATABASE', 'DASY-FLII-TEST'),
            username=os.getenv('TEST_SQL_USERNAME', 'test_user'),
            password=os.getenv('TEST_SQL_PASSWORD', 'test_password'),
        ),
        Environment.PRODUCTION: DatabaseConfig(
            server=os.getenv('PROD_SQL_SERVER', 'prod_server'),
            database=os.getenv('PROD_SQL_DATABASE', 'DASY-FLII'),
            username=os.getenv('PROD_SQL_USERNAME', 'prod_user'),
            password=os.getenv('PROD_SQL_PASSWORD', 'prod_password'),
        ),
    })

    # AppSheet configuration
    appsheet: AppSheetConfig = field(default_factory=lambda: AppSheetConfig(
        app_id=os.getenv('APPSHEET_APP_ID', 'your_app_id'),
        api_key=os.getenv('APPSHEET_API_KEY', 'your_api_key'),
        table_name=os.getenv('APPSHEET_TABLE_NAME', 'PRLocalPurchase'),
    ))

    # Document search settings
    auto_no_chr: str = field(default_factory=lambda: os.getenv('AUTO_NO_CHR', 'PRLOC'))
    auto_no: str = field(default_factory=lambda: os.getenv('AUTO_NO', ''))  # For recovery mode

    @property
    def db(self) -> DatabaseConfig:
        """Get current database configuration based on environment"""
        return self.db_configs[self.environment]

    def __post_init__(self):
        logger.info(f"Configuration loaded for environment: {self.environment.value}")
        logger.info(f"Database: {self.db.database} on {self.db.server}")


# ===================== DATA MODELS =====================

@dataclass
class PRDocument:
    """PR Local Purchase document data"""
    sinsei_code: str
    company_code: str = ""      # จาก SINSEI_SYOZOKU_CODE → AppSheet: company
    document_id: str = ""       # จาก AUTO_NO → AppSheet: document_running_no
    links: List[str] = field(default_factory=list)  # Links จาก TR_SINSEI_DATA_LNK (LNK1-LNK5)

    def to_appsheet_rows(self) -> List[Dict[str, Any]]:
        """Convert to AppSheet rows format (1 row per link)"""
        rows = []
        for link in self.links:
            if link:  # เฉพาะ link ที่ไม่เป็น NULL/empty
                rows.append({
                    "id_random_download": self.sinsei_code,
                    "company": self.company_code,
                    "document_running_no": self.document_id,
                    "link": link,
                })
        return rows


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

    def get_approved_documents(self, limit: int = 1) -> List[PRDocument]:
        """
        Query to get approved PR Local Purchase documents

        Args:
            limit: Maximum number of documents to retrieve

        Returns:
            List of PRDocument objects
        """
        if self.config.auto_no:
            # Recovery mode: Search by exact document number
            query = f"""
                SELECT TOP {limit} h.SINSEI_CODE
                FROM FLIISA.TR_SINSEI_DATA_HEADER h
                WHERE h.AUTO_NO = ?
            """
            params = (self.config.auto_no,)
            logger.info(f"Recovery mode: Searching for AUTO_NO = '{self.config.auto_no}'")
        else:
            # Normal mode: Search for approved documents
            query = f"""
                SELECT TOP {limit} h.SINSEI_CODE
                FROM FLIISA.TR_SINSEI_DATA_HEADER h
                WHERE h.AUTO_NO_CHR = ?
                    AND h.JOUTAI_KBN = 1
                    AND DATEPART(MILLISECOND, h.INS_DATE) != 0
                    AND DATEPART(MILLISECOND, h.UPD_DATE) != 0
                    AND DATEPART(SECOND, h.INS_DATE) != 0
                    AND DATEPART(SECOND, h.UPD_DATE) != 0
                ORDER BY h.UPD_DATE ASC
            """
            params = (self.config.auto_no_chr,)
            logger.info(f"Normal mode: Searching for AUTO_NO_CHR = '{self.config.auto_no_chr}'")

        try:
            cursor = self.connection.cursor()
            cursor.execute(query, params)
            rows = cursor.fetchall()

            documents = []
            for row in rows:
                doc = PRDocument(
                    sinsei_code=row[0],
                    # google_drive_link และ company_code จะถูก set ใน process_document()
                )
                documents.append(doc)

            logger.info(f"Found {len(documents)} approved document(s)")
            return documents

        except pyodbc.Error as e:
            logger.error(f"Error executing query: {e}")
            raise

    def get_google_drive_link(self, sinsei_code: str) -> Optional[str]:
        """
        Get Google Drive link from TR_SINSEI_DATA_LNK

        Args:
            sinsei_code: Document SINSEI_CODE

        Returns:
            Google Drive link or None
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
                logger.info(f"Found Google Drive link for {sinsei_code}")
                return row[0]
            else:
                logger.warning(f"No Google Drive link found for {sinsei_code}")
                return None

        except pyodbc.Error as e:
            logger.error(f"Error getting Google Drive link: {e}")
            raise

    def get_all_links(self, sinsei_code: str) -> List[str]:
        """
        Get all Google Drive links (LNK1-LNK5) from TR_SINSEI_DATA_LNK

        Args:
            sinsei_code: Document SINSEI_CODE

        Returns:
            List of Google Drive links (excluding NULL values)
        """
        query = """
            SELECT KOUMOKU_VALUE
            FROM FLIISA.TR_SINSEI_DATA_LNK
            WHERE SINSEI_CODE = ?
                AND KOUMOKU_KEY IN ('LNK1', 'LNK2', 'LNK3', 'LNK4', 'LNK5')
            ORDER BY KOUMOKU_KEY
        """

        try:
            cursor = self.connection.cursor()
            cursor.execute(query, (sinsei_code,))
            rows = cursor.fetchall()

            # Filter out NULL values
            links = [row[0] for row in rows if row[0]]
            logger.info(f"Found {len(links)} link(s) for {sinsei_code}")
            return links

        except pyodbc.Error as e:
            logger.error(f"Error getting links: {e}")
            raise

    def get_document_details(self, sinsei_code: str) -> Dict[str, Any]:
        """
        Get document details from TR_SINSEI_DATA_HEADER

        Args:
            sinsei_code: Document SINSEI_CODE

        Returns:
            Dictionary with document details:
            - COMPANY_CODE: from SINSEI_SYOZOKU_CODE
            - DOCUMENT_ID: from AUTO_NO
        """
        query = """
            SELECT SINSEI_SYOZOKU_CODE, AUTO_NO
            FROM FLIISA.TR_SINSEI_DATA_HEADER
            WHERE SINSEI_CODE = ?
        """

        try:
            cursor = self.connection.cursor()
            cursor.execute(query, (sinsei_code,))
            row = cursor.fetchone()

            details = {}
            if row:
                details["COMPANY_CODE"] = row[0] or ""
                details["DOCUMENT_ID"] = row[1] or ""

            logger.info(f"Retrieved details for {sinsei_code}: company_code={details.get('COMPANY_CODE')}, document_id={details.get('DOCUMENT_ID')}")
            return details

        except pyodbc.Error as e:
            logger.error(f"Error getting document details: {e}")
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
        # Update INS_DATE and UPD_DATE to remove milliseconds and seconds
        # This marks the document as "processed"
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
            logger.error(f"Response text was: {response.text[:500] if response.text else '(empty)'}")
            return {"success": False, "error": f"Invalid JSON response: {e}"}

    def add_rows(self, rows: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Add multiple rows to AppSheet table

        Args:
            rows: List of row data dictionaries

        Returns:
            API response
        """
        payload = {
            "Action": "Add",
            "Properties": {
                "Locale": "en-US",
                "Timezone": "SE Asia Standard Time"
            },
            "Rows": rows
        }

        try:
            logger.info(f"Adding {len(rows)} row(s) to AppSheet: {self.app_config.table_name}")
            logger.info(f"API URL: {self.api_url}")
            logger.info(f"Payload: {payload}")

            response = requests.post(
                self.api_url,
                json=payload,
                headers=self.headers,
                timeout=60
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
            logger.info(f"Successfully added {len(rows)} row(s) to AppSheet")
            return {"success": True, "response": result}

        except requests.exceptions.RequestException as e:
            logger.error(f"Error adding rows to AppSheet: {e}")
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
        # Use Find action to test connection
        payload = {
            "Action": "Find",
            "Properties": {
                "Locale": "th-TH"
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

class PRLocalPurchaseService:
    """Main service for processing PR Local Purchase documents"""

    def __init__(self, config: Config):
        self.config = config
        self.sql_client = SQLServerClient(config)
        self.appsheet_client = AppSheetClient(config)

    def process_document(self, doc: PRDocument) -> Dict[str, Any]:
        """
        Process a single document

        Args:
            doc: PRDocument to process

        Returns:
            Processing result
        """
        result = {
            "sinsei_code": doc.sinsei_code,
            "success": False,
            "steps": []
        }

        try:
            # Step 1: Get document details from TR_SINSEI_DATA_HEADER
            details = self.sql_client.get_document_details(doc.sinsei_code)
            doc.company_code = details.get("COMPANY_CODE", "")   # SINSEI_SYOZOKU_CODE
            doc.document_id = details.get("DOCUMENT_ID", "")     # AUTO_NO
            result["steps"].append({"step": "get_details", "success": True})

            # Step 2: Get all links from TR_SINSEI_DATA_LNK
            doc.links = self.sql_client.get_all_links(doc.sinsei_code)
            result["steps"].append({"step": "get_links", "success": True, "count": len(doc.links)})

            # Step 3: Add to AppSheet (1 row per link)
            rows = doc.to_appsheet_rows()
            if not rows:
                result["steps"].append({
                    "step": "add_to_appsheet",
                    "success": False,
                    "error": "No links found to insert"
                })
                return result

            appsheet_result = self.appsheet_client.add_rows(rows)
            if appsheet_result["success"]:
                result["steps"].append({"step": "add_to_appsheet", "success": True, "rows_added": len(rows)})

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
        logger.info(f"Starting PR Local Purchase Processor")
        logger.info(f"Environment: {self.config.environment.value}")
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
                logger.info(f"Processing document: {doc.sinsei_code}")
                result = self.process_document(doc)
                results.append(result)

                if result["success"]:
                    logger.info(f"Document {doc.sinsei_code} processed successfully")
                else:
                    logger.warning(f"Document {doc.sinsei_code} processing failed")

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
    service = PRLocalPurchaseService(config)

    try:
        results = service.run(limit=1)

        print("\n" + "=" * 60)
        print("Processing Results:")
        print("=" * 60)

        for result in results:
            status = "SUCCESS" if result["success"] else "FAILED"
            print(f"\n  SINSEI_CODE: {result['sinsei_code']} - {status}")
            for step in result.get("steps", []):
                step_status = "OK" if step["success"] else "FAIL"
                print(f"    - {step['step']}: {step_status}")
                if "error" in step:
                    print(f"      Error: {step['error']}")

        successful = sum(1 for r in results if r["success"])
        print(f"\nTotal: {successful}/{len(results)} document(s) processed successfully")

    except Exception as e:
        logger.error(f"Process failed: {e}")
        raise


if __name__ == "__main__":
    main()
