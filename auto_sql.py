import logging
from logging.handlers import TimedRotatingFileHandler
import os
from datetime import datetime

import cx_Oracle  # Oracle connection
import pandas as pd  # Data analysis


from config import (
    DB_CONFIG,
    EXCEL_ENGINE,
    LOG_DIRECTORY,
    LOG_FILE,
    LOG_FORMAT,
    LOG_LEVEL,
    LOG_ROTATION_INTERVAL,
    LOG_ROTATION_WHEN,
    LOG_BACKUP_COUNT,
    ORACLE_CLIENT_PATH,
    OUTPUT_DIRECTORY,
    OUTPUT_FILENAME_TEMPLATE,
)


class ApplicationError(Exception):
    """Base class for all application-specific exceptions."""
    pass


class DatabaseConnectionError(ApplicationError):
    """Raised when the Oracle database connection fails."""
    pass


class QueryExecutionError(ApplicationError):
    """Raised when an Oracle SQL query fails."""
    pass


class DataFrameBuildError(ApplicationError):
    """Raised when building the DataFrame fails."""
    pass


class ReportExportError(ApplicationError):
    """Raised when exporting the Excel report fails."""
    pass

PO_DISTRIBUTION_ISSUE_SQL = """
        SELECT
            PO_DISTRIBUTION_ID,
            PO_HEADER_ID,
            PO_LINE_ID,
            REQ_DISTRIBUTION_ID,
            DELIVER_TO_LOCATION_ID,
            DELIVER_TO_PERSON_ID,
            CREATION_DATE
        FROM PO_DISTRIBUTIONS_ALL
        WHERE REQ_DISTRIBUTION_ID IS NULL
            AND DELIVER_TO_LOCATION_ID IS NULL
            AND DELIVER_TO_PERSON_ID IS NULL
        ORDER BY CREATION_DATE DESC
        """
def setup_logger():
    """Configure and return the application logger."""

    os.makedirs(
        LOG_DIRECTORY,
        exist_ok=True,
    )

    logger = logging.getLogger(__name__)
    logger.setLevel(LOG_LEVEL)
    logger.propagate = False

    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        LOG_FORMAT
    )

    file_path = os.path.join(
        LOG_DIRECTORY,
        LOG_FILE,
    )

    file_handler = TimedRotatingFileHandler(
    file_path,
    when=LOG_ROTATION_WHEN,
    interval=LOG_ROTATION_INTERVAL,
    backupCount=LOG_BACKUP_COUNT,
    encoding="utf-8",
    )
    file_handler.setLevel(LOG_LEVEL)
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(LOG_LEVEL)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

logger = setup_logger()

_oracle_client_initialized = False


def initialize_oracle_client():
    """Load the Oracle Instant Client selected in .env.

    This prevents cx_Oracle from accidentally loading an incompatible oci.dll
    found earlier on PATH (for example, a legacy 32-bit Oracle installation).
    """
    global _oracle_client_initialized

    if _oracle_client_initialized:
        return

    client_path = os.path.abspath(os.path.expandvars(ORACLE_CLIENT_PATH))
    oci_dll = os.path.join(client_path, "oci.dll")

    if not os.path.isdir(client_path):
        raise DatabaseConnectionError(
            "Oracle Client directory does not exist: "
            f"{client_path}. Check ORACLE_CLIENT_PATH in .env."
        )

    if os.name == "nt" and not os.path.isfile(oci_dll):
        raise DatabaseConnectionError(
            f"Oracle Client library was not found: {oci_dll}"
        )

    try:
        cx_Oracle.init_oracle_client(lib_dir=client_path)
        _oracle_client_initialized = True
        logger.info("Oracle Client loaded from: %s", client_path)
    except cx_Oracle.Error as error:
        raise DatabaseConnectionError(
            "Unable to load Oracle Client from "
            f"{client_path}: {error}. Make sure Python and Oracle Instant "
            "Client are both 64-bit (or both 32-bit)."
        ) from error

def create_connection():
    """Create and return an Oracle database connection."""
    try:
        initialize_oracle_client()
        return cx_Oracle.connect(
            user=DB_CONFIG["account"],
            password=DB_CONFIG["password"],
            dsn=DB_CONFIG["dsn"],
            encoding=DB_CONFIG["encoding"],
        )

    except cx_Oracle.Error as error:
        raise DatabaseConnectionError(
            f"Unable to connect to Oracle database: {error}"
        ) from error


def fetch_query_results(connection, query):
    """Execute an Oracle query and return rows and column names."""

    try:
        with connection.cursor() as cursor:
            cursor.execute(query)

            column_names = [column[0] for column in cursor.description]
            rows = cursor.fetchall()

            return rows, column_names

    except cx_Oracle.Error as error:
        raise QueryExecutionError(
            f"Failed to execute Oracle query: {error}"
        ) from error
    
def build_dataframe(rows,column_names):
    try:
    # ===== 轉成 DataFrame =====
        df = pd.DataFrame(rows, columns=column_names)
        return df
    except (KeyError, ValueError) as error:
        raise DataFrameBuildError(
            f"Failed to build DataFrame: {error}"
        ) from error

def export_excel(dataframe):
    """Export the DataFrame to a dated Excel report."""
    try:
        os.makedirs(OUTPUT_DIRECTORY, exist_ok=True)
        today = datetime.now().strftime("%Y%m%d")

        output_filename = OUTPUT_FILENAME_TEMPLATE.format(
            date=today
        )

        file_name = os.path.join(
            OUTPUT_DIRECTORY,
            output_filename,
        )

        dataframe.to_excel(
            file_name,
            index=False,
            engine=EXCEL_ENGINE,
        )

        return file_name

    except OSError as error:
        raise ReportExportError(
            f"Failed to export Excel report: {error}"
        ) from error
        
def close_connection(connection):
    """Close the Oracle database connection if it exists."""
    if connection is not None:
        connection.close()
        logger.info("Oracle database connection closed.")

def run_po_distribution_check():
    connection = None

    try:
        connection = create_connection()
        logger.info("Oracle database connection established.")

        rows, columns = fetch_query_results(
            connection,
            PO_DISTRIBUTION_ISSUE_SQL,
        )

        issue_count = len(rows)
        #print(f"資料品質異常筆數：{issue_count}")
        
        logger.info(
            "Data quality issue count: %s",
            issue_count,
        )

        if not rows:
            #print("未發現資料品質問題")
            logger.info(
                "No data quality issues found."
            )
            return

        dataframe = build_dataframe(rows, columns)
        file_name = export_excel(dataframe)

        #print(f"已輸出資料品質報告：{file_name}")
        logger.info(
            "Data quality report exported: %s",
            file_name,
        )

    except DatabaseConnectionError as error:
        logger.error(error)

    except QueryExecutionError as error:
        logger.error(error)

    except DataFrameBuildError as error:
        logger.error(error)

    except ReportExportError as error:
        logger.error(error)
        
    finally:
        close_connection(connection)
        
def oracle():
    output = input("請選擇輸出格式：")
    output = output.lower().strip();
    if output not in ['excel','csv']: #錯誤處理
     print("請輸入excel/csv")
     return#避免繼續執行
 
    # ===== 連線資料 =====
    account = "APPS"#oracle使用者
    pwd = "APPS"#密碼
    dsn = "192.168.100.43:1541/C2504"#IP:PORT/SID

    sqlPo = """
    SELECT
        PO_DISTRIBUTION_ID,
        PO_HEADER_ID,
        PO_LINE_ID,
        REQ_DISTRIBUTION_ID,
        DELIVER_TO_LOCATION_ID,
        DELIVER_TO_PERSON_ID,
        CREATION_DATE
    FROM PO_DISTRIBUTIONS_ALL
    WHERE REQ_DISTRIBUTION_ID IS NULL
      AND DELIVER_TO_LOCATION_ID IS NULL
      AND DELIVER_TO_PERSON_ID IS NULL
    ORDER BY CREATION_DATE DESC
    """

    try:#防止連線失敗
        # ===== 連線 Oracle =====
        con = cx_Oracle.connect(account, pwd, dsn, encoding="UTF-8")
        cursor = con.cursor()
        # cursor用於執行Sql、獲取查詢結果

        # ===== 執行 SQL =====
        cursor.execute(sqlPo)
        result = cursor.fetchall()
        # fetchone()    fetchmany(10)
        issue_cnt = len(result)

        print(f"資料品質異常筆數：{issue_cnt}")

        if issue_cnt > 0:
            # ===== 轉成 DataFrame =====
            columns = [col[0] for col in cursor.description]
            # col[0]:name - Column name
            # col[1]:type_code - Column type code
            # col[2]:display_size - Display size
            # col[3]:internal_size - Internal size
            # col[4]:precision - Precision
            # col[5]:scale - Scale
            # col[6]:null_ok - Whether NULL values are allowed
            # col[7]:field_flags - Field flags (extension to PEP-249)
            # col[8]:table_name - Table name (extension to PEP-249)
            # col[9]:original_column_name - Original column name (extension to PEP-249)
            # col[10]:original_table_name - Original table name (extension to PEP-249)
            print(f"資料品質：{columns}")
            df = pd.DataFrame(result, columns=columns)

            today = datetime.now().strftime("%Y%m%d")
            # file_name = f"DQ_PO_DISTRIBUTIONS_{today}.xlsx"
            file_name = f"DQ_PO_DISTRIBUTIONS_{today}."

            if output == 'excel':
             file_name += ".xlsx"
             df.to_excel(
                 file_name,
                 index=False,#不輸出 pandas index
                 engine="openpyxl"#Excel 引擎
             )
                        # 在 df.to_excel() 前加入：
            elif output == 'csv':
             file_name += ".csv"
             df.to_csv(
                 file_name,
                 index=False,#不輸出 pandas index
                 encoding='utf-8-sig'
             )
            if df is not None:
             dataFrame_try(df)
            print(f"已輸出資料品質報告：{file_name}")
        else:
            print("未發現資料品質問題 🎉")

    except Exception as e:
        print("程式發生錯誤：", e)

    finally:
        if cursor: cursor.close()#關閉前先確認存在
        if con: con.close()#關閉前先確認存在
        print("資料庫連線已關閉")
        
def dataFrame_try(df):
    print("=== 資料品質檢查 ===")
    print(f"總異常筆數：{df.shape[0]}")
    print(df.groupby('PO_HEADER_ID').size().head())  # PO_HEADER_ID順序前5組 head()預設取前5
    
    print("\n欄位型態：")
    print(df.dtypes)
    
    print("\n前3筆：")
    print(df.head(3))
    
    print("\n後3筆：")
    print(df.tail(3))
    
    print("\nPO_HEADER_ID分佈：")
    print(df.describe())
    
    print("\n下SQL：")
    print(df.query("REQ_DISTRIBUTION_ID.isna()"))#REQ_DISTRIBUTION_ID is NULL
    print(df.groupby('PO_HEADER_ID').size())#groupby分組
    
    print("\n=== Top5異常PO ===")
    top5_po = df.groupby('PO_HEADER_ID').size().nlargest(5)
    print(top5_po)
    print(f"\n最多異常PO：{top5_po.index[0]} (異常{top5_po.iloc[0]}筆)")


if __name__ == "__main__":
     run_po_distribution_check()

