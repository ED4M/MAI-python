import pandas as pd
import numpy as np
import cx_Oracle
import os
import config

print('Reading csv-file...')
# cp1251
df = pd.read_csv("import.csv",sep=';',encoding="cp1251") 

mapped = df[(df['ADDRESS'] != None) & (df['ADDRESS'].str.contains("Москва"))]
mapped = mapped.replace(np.nan, '', regex=True)

print('Writing data to Excel-file...')
writer = pd.ExcelWriter('export.xlsx', engine='xlsxwriter')
mapped.to_excel(writer, 'Sheet1')
writer.save()

# Oracle
LOCATION_ORACLE = r"D:\Prog\instantclient_19_9"
os.environ["PATH"] = LOCATION_ORACLE + ";" + os.environ["PATH"]

conn = cx_Oracle.connect(
    config.user, 
    config.password, 
    config.dsn, 
    encoding=config.encoding)
curs = conn.cursor()

print('Exporting data to Oracle DB...')
for item in mapped.values:
  curs.execute(
    "INSERT INTO FNS_BASE (KOD, NAIMK, NAIM_FULL, ADDRESS, COMMENTS) VALUES (:KOD, :NAIMK, :NAIM_FULL, :ADDRESS, :COMMENTS)",
      KOD=item[0], NAIMK=item[1], NAIM_FULL=item[2], ADDRESS=item[3], COMMENTS=item[4]
      )

conn.commit()
print('Done importing!')
conn.close()
