import pandas as pd

class ExcelToSQL:
    def __init__(self,excel_file_path):
        # 讀取 Excel 文件
        self.excel_file_path = excel_file_path
    
    def create_dataframe(self,excel_file_path):
        df = pd.read_excel(excel_file_path, dtype=str)
        return df
    
    def get_dataframe_columns(self,df):
        # 提取列名稱（Header）
        columns = df.columns.tolist()
        return columns

    def create_insert_sql(self,df,columns):
        # 迭代 DataFrame 的每一行，生成插入語句
        with open('PDATA_PUB_DEPT_HIST.sql', 'w',encoding='utf-8') as sql_file:
            for index, row in df.iterrows():
                values = ', '.join([f"'{value}'" if isinstance(value, str) else str(value) for value in row])
                
                # 生成 INSERT INTO 語法
                sql_insert = f"INSERT INTO PDATA.PUB_DEPT_HIST ({', '.join(columns)}) VALUES ({values});"
                sql_insert = sql_insert.replace('nan', 'NULL')
                
                # 打印或保存 SQL Insert 語法
                print(sql_insert)
                sql_file.write(sql_insert + '\n')
        return True

    def excel_to_sql_process(self):
        df = self.create_dataframe(self.excel_file_path)
        columns = self.get_dataframe_columns(df)
        self.create_insert_sql(df,columns)
        return True

if __name__ == "__main__":
    excel_to_sql = ExcelToSQL('PDATA_PUB_DEPT_HIST.xlsx')
    excel_to_sql.excel_to_sql_process()