import mysql.connector, MySQLdb, xlrd, xlwt,os

# detecting xls file
dir_path = os.path.dirname(os.path.realpath(__file__))
all_files = os.listdir(dir_path+"/Excel Files/")
if len(all_files)>1:
    print "More than one file detected in the 'Excel Files' folder. Use different worksheets to create multiple tables. A single .xls file represents a database."
    exit()

#importing xls file
excel_file=dir_path+"/Excel Files/"+all_files[0]
length_front_cutoff= len(dir_path+"/Excel Files/")
length_back_cutoff=len(".xls")

#set up database name
db_name=excel_file[length_front_cutoff:len(excel_file)-4]

#import workbook and sheets
workbook = xlrd.open_workbook(excel_file)
number_of_tables=len(workbook.sheet_names())

#create connection to mysql
mydb=mysql.connector.connect(host="localhost",user="root")

#init cursor
mycursor=mydb.cursor()


#delete database
create_statement = "DROP DATABASE {:s}".format(db_name)
mycursor.execute(create_statement)

#create and select database
create_statement = "CREATE DATABASE {:s}".format(db_name)
mycursor.execute(create_statement)


#select database
create_statement = "USE {:s}".format(db_name)
mycursor.execute(create_statement)

#count nr of sheets in workbook
number_of_sheets = len(workbook.sheet_names()) 


# HALLO??

######################################################################################################################################################################################################################################################################################################################################################
'''JUST 1 TABLE'''
worksheet = workbook.sheet_by_index(0)

#counting number_of_rows and columns
number_of_rows=worksheet.nrows-1
number_of_columns=worksheet.ncols

#create table
mycursor.execute("CREATE TABLE Test (id MEDIUMINT NOT NULL AUTO_INCREMENT,PRIMARY KEY (id));")

#identify datatype and name of column, then create column in sql
for i in range(number_of_columns):
    
    column_title=worksheet.cell(0,i).value
    datatype_of_column=type(worksheet.cell(1,i).value)

    if datatype_of_column==float:
        cell_type="FLOAT(10)"
    elif datatype_of_column==int:
        cell_type="INT(10)"
    elif datatype_of_column==unicode or datatype_of_column==str:
        cell_type="VARCHAR(100)"
    elif datatype_of_column==bool:
        cell_type="BINARY"
    else:
        print ('Datatype at column %s not recognized. Set to float.', i)
    
    mycursor.execute("ALTER TABLE Test ADD (%s %s)" % (column_title, cell_type))

'''
#insert values to columns
for i in range(number_of_columns):
    column_title=worksheet.cell(0,i).value

    for n in range(1,number_of_rows+1):

        value_for_table=worksheet.cell(n,i).value
        print n,i,column_title, value_for_table

        query=("INSERT INTO Test (%s)" % column_title) + " VALUES (%s)"
        mycursor.execute(query, (value_for_table,))
        #mycursor.execute("INSERT INTO Test (%s) VALUES (%s)", (column_title,value_for_table))
    '''

for n in range(1,number_of_rows+1):

    value_for_table=worksheet.cell(n,i).value
    query="INSERT INTO Test ("
    query_values="("
    values=[]
    for i in range(number_of_columns):
        column_title=worksheet.cell(0,i).value
        value_for_table = worksheet.cell(n, i).value
        query=query+column_title if i==0 else query+','+column_title
        query_values=query_values+"%s" if i==0 else query_values+',%s'
        values.append(value_for_table)
    query=query+') VALUES '+query_values+')'
    mycursor.execute(query, values)

mydb.commit()
mydb.close()

