import cx_Oracle#oracle連線
import pandas as pd#資料分析
from datetime import datetime#取得系統時間

# Oracle Instant Client
cx_Oracle.init_oracle_client(lib_dir=r"D:\instantclient_21_7")#初始化 Oracle Instant Client


def oracle():
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

        # ===== 執行 SQL =====
        cursor.execute(sqlPo)
        result = cursor.fetchall()
        issue_cnt = len(result)

        print(f"資料品質異常筆數：{issue_cnt}")

        if issue_cnt > 0:
            # ===== 轉成 DataFrame =====
            columns = [col[0] for col in cursor.description]
            df = pd.DataFrame(result, columns=columns)

            today = datetime.now().strftime("%Y%m%d")
            file_name = f"DQ_PO_DISTRIBUTIONS_{today}.xlsx"

            df.to_excel(
                file_name,
                index=False,#不輸出 pandas index
                engine="openpyxl"#Excel 引擎
            )

            print(f"已輸出資料品質報告：{file_name}")
        else:
            print("未發現資料品質問題 🎉")

    except Exception as e:
        print("程式發生錯誤：", e)

    finally:
        cursor.close()
        con.close()
        print("資料庫連線已關閉")


if __name__ == "__main__":
    oracle()
