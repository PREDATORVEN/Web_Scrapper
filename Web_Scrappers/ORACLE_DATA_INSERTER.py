import csv
import os
import glob
import cx_Oracle
from datetime import datetime

# Oracle database connection details
username = 'sys'
password = 'root'
service_name = "orcl"
dsn = cx_Oracle.makedsn("localhost", 1522, service_name=service_name)
#------------------date module----
today_date=datetime.now().strftime("%d-%b-%y").upper()
print(today_date)
#---------------------------------
# CSV file path
downloads_folder = r"D:\YASH STUFF\NET_SCHEDULE_FILE"
if not os.path.isdir(downloads_folder):
    raise ValueError("The provided path is not a valid directory.")
csv_files = glob.glob(os.path.join(downloads_folder, '*.csv'))

# Check if any CSV files were found
if not csv_files:
    print(" none")  
# Sort CSV files by modification time in descending order
csv_files.sort(key=os.path.getmtime, reverse=True)
latest_csv_file =csv_files[0]
print(f"The latest downloaded CSV file is: {latest_csv_file}")


#---------------------------------
# csv_file_path = r"C:\Users\Yash Sharma\OneDrive\Desktop\NetSchedule-Summary-ALL_Seller(135)-18-08-2023.csv"
csv_file_path =latest_csv_file
# csv_file_path=r"D:\YASH STUFF\Python\file.csv"
# List of Header_Name values
with open(csv_file_path, 'r') as file:
    csv_reader = csv.reader(file)
    print(type(csv_reader))
    print("Headers",csv_reader)
    header = next(csv_reader)
sum=0
data=[]
header_names=[]
for column_name in header:
    if column_name not in data:
        data.append(column_name)
        print("column_name: ",column_name)
        sum+=1
    # collection_schema[column_name] = {'type': 'string'}
print("total columns: ",sum)
file.close
tab=0
print("TABLE TO BE MAID ARE: ")
for i in data:
    if i not in ['Time Block','Time Desc','Grand Total']:
        header_names.append(i)

# Establish database connection
connection = cx_Oracle.connect(username, password, dsn, mode=cx_Oracle.SYSDBA)
cursor = connection.cursor()
header_names_len=len(header_names)
i=0
row_index=2
TAB_CNT=1
print("length of list",header_names_len)
while(i<header_names_len):
    with open(csv_file_path, "r") as file:
        csv_reader = csv.reader(file)
        next(csv_reader)  # Skip the first two rows
        next(csv_reader)

        for row in csv_reader:
            schedule_id = row[0]
            time_slot = row[1]
            row_data = row[row_index:row_index+14]  # Extract data for the current header

            placeholders = ', '.join([':' + str(i) for i in range(1, len(row_data) + 1)])
            insert_query = f'''
                INSERT INTO NET_SCHEDULE(SCHEDULE_ID,ENTRY_DT, TIME_SLOT, Header_Name, ISGS_VAL, MTOA_VAL, STOA_VAL, LTA_VAL, DAM_VAL, GDAM_VAL, URS_VAL, AS_VAL, SCED_VAL, REMC_VAL, RTM_VAL, HPDAM_VAL, BUNDLE_RE,TOTAL_VAL) 
                VALUES (:1,TO_DATE(:2, 'DD-MON-YY'), :3,:4, {placeholders})
            '''
            cursor.execute(insert_query, [schedule_id,today_date,time_slot, header_names[i]] + row_data)

        print("Data inserted successfully.")
        print("table number",TAB_CNT)
        print("value of index i ",i)
    i+=1
    row_index+=14
    TAB_CNT+=1    

# Commit and close the connection
connection.commit()
cursor.close()
connection.close()
