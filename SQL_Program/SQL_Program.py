from typing import List, Dict, Any
import pandas as pd
import sqlalchemy as sa
from sqlalchemy.orm import sessionmaker, declarative_base
from sqlalchemy.exc import SQLAlchemyError
import glob
import sys
import os
from datetime import date
from contextlib import contextmanager

# 自作関数の定義
sys.path.append('../MyModule')
import Log

# ----- SQLサーバへの接続情報 -----
cle_conf = {
    "ID": "readonly",
    "PASS": "Re0d0n1y",
    "HOST": "tdsprd08.c0uzg0vwy8aj.ap-northeast-1.rds.amazonaws.com",
    "SID": "TDSPRD08",
    "PORT": 1525
}

# SQLAlchemy setup
Base = declarative_base()
connection_string = f'oracle+cx_oracle://{cle_conf["ID"]}:{cle_conf["PASS"]}@{cle_conf["HOST"]}:{cle_conf["PORT"]}/?service_name={cle_conf["SID"]}'
engine = sa.create_engine(connection_string)
Session = sessionmaker(bind=engine)

@contextmanager
def db_session():
    """Provide a transactional scope around a series of operations."""
    session = Session()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()

def setup_logging() -> str:
    """Setup logging directory and return log file path"""
    Log_Folder_Name = "-".join(str(date.today()).split("-"))
    log_dir = os.path.join("..", "Log", Log_Folder_Name)
    
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    return os.path.join(log_dir, 'SQL_Program.log')

def execute_sql_files(sql_files: List[str], log_file: str) -> None:
    """Execute SQL files and save results to Excel and CSV"""
    for sql_file in sql_files:
        try:
            with open(sql_file, 'r') as f:
                sql_query = f.read()
            
            with db_session() as session:
                result = session.execute(sa.text(sql_query))
                df = pd.DataFrame(result.fetchall(), columns=result.keys())

                # Save to Excel and CSV
                excel_path = os.path.join('C:/Users/hsi67063/Downloads/Python test/excel/', 
                                        sql_file.replace('.sql', '.xlsx'))
                csv_path = os.path.join('C:/Users/hsi67063/Downloads/Python test/csv/',
                                      sql_file.replace('.sql', '.csv'))
                
                df.to_excel(excel_path, sheet_name='Sheet1', index=False)
                df.to_csv(csv_path, index=False, encoding='cp932')
                
                Log.Log_Info(log_file, f"{sql_file} : OK")
                
        except SQLAlchemyError as e:
            Log.Log_Info(log_file, f"{sql_file} : Error - {str(e)}")
            continue

def main():
    # Change to SQL directory
    os.chdir('../SQL2/')
    sql_files = glob.glob('*.sql')
    
    # Setup logging
    log_file = setup_logging()
    Log.Log_Info(log_file, "Program Start")
    
    # Execute SQL files
    execute_sql_files(sql_files, log_file)
    
    Log.Log_Info(log_file, "Program End")

if __name__ == "__main__":
    main()
