# USED FOR DB ACCESS
import sqlalchemy

# PYTHON DATA ANALYSIS LIBRARY
import pandas as pd

df = pd.read_excel('VDF_Cells.xlsx', sheet_name='Tabelle1')

# WRITE TO VOICE
engine_v = sqlalchemy.create_engine(f'mssql+pyodbc://blndb11/DE_BM_Voice_1907?driver=SQL+Server+Native+Client+11.0')
df.to_sql('VDF_DE_4G_Cells', con=engine_v, if_exists='replace')

# WRITE TO DATA
engine_d = sqlalchemy.create_engine(f'mssql+pyodbc://blndb11/DE_BM_Data_1907?driver=SQL+Server+Native+Client+11.0')
df.to_sql('VDF_DE_4G_Cells', con=engine_d, if_exists='replace')