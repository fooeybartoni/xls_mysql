import pymysql.cursors
import xlrd
 
#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
 
    # print number of sheets
    print book.nsheets
 
    # print sheet names
    print book.sheet_names()
 
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
 
    # read a row
    print first_sheet.row_values(0)
 
    # read a cell
    cell = first_sheet.cell(0,0)
    print cell
    print cell.value
 
    # read a row slice
    print first_sheet.row_slice(rowx=0,
                                start_colx=0,
                                end_colx=2)
 
#----------------------------------------------------------------------
def DBConnect(host,user,passwd,db)
    # Connect to the database
    connection = pymysql.connect(host='localhost',
                                user='root',
                                passwd='S3cur1ty!',
                                db='ita_cnd',
                                charset='utf8mb4',
                                cursorclass=pymysql.cursors.DictCursor)
    
    try:
        with connection.cursor() as cursor:
            # Create a new record
            sql = "INSERT INTO `users` (`email`, `password`) VALUES (%s, %s)"
            cursor.execute(sql, ('webmaster@python.org', 'very-secret'))
    
        # connection is not autocommit by default. So you must commit to save
        # your changes.
        connection.commit()
    
        with connection.cursor() as cursor:
            # Read a single record
            sql = "SELECT `id`, `password` FROM `users` WHERE `email`=%s"
            cursor.execute(sql, ('webmaster@python.org',))
            result = cursor.fetchone()
            print(result)
    finally:
        connection.close()
        
#----------------------------------------------------------------------
if __name__ == "__main__":
    path = "/home/cnd_user/dev/53.xlsx"
    open_file(path)
