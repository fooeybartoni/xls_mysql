import pymysql.cursors
import xlrd
 
#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    rowList = list()
    workbook = xlrd.open_workbook(path)
    worksheet = workbook.sheet_by_name('Rev4')
    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1
    curr_row = -1
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        #print 'Row:', curr_row
        curr_cell = -1
        cellList = []
        while curr_cell < num_cells:
            curr_cell += 1
            # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
            cell_type = worksheet.cell_type(curr_row, curr_cell)
            cell_value = worksheet.cell_value(curr_row, curr_cell)
            #print '	', cell_type, ':', cell_value
            cellList.append(cell_value)
        rowList.append(cellList)
    return rowList
#----------------------------------------------------------------------
def DBConnect(host,user,passwd,db):
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
    testList = open_file(path)
    row_cnt = 0
    out_str = ''
    for row in testList:
        row_cnt += 1
        out_str += "\nRow " + str(row_cnt) + "\n" 
        col_cnt = 0 
        for cellVal in row:
            col_cnt+=1
            inVal = cellVal
            if cellVal and cellVal.strip():
                inVal = cellVal
            else:
                inVal = "***Blank No Data***"
            out_str += (("\tColumn " + str(col_cnt) +" -- "+inVal+"\n"))
    print out_str
    filename = "/home/cnd_user/dev/test_data1.txt"    
    target = open(filename, 'w')
    target.write(out_str.encode('utf8'))
    target.close()