import xlrd
import MySQLdb
from auth import AUTH

# Open the workbook and define the worksheet
book = xlrd.open_workbook("2016.xlsx")
sheet = book.sheet_by_name("Sheet1")

# Establish a MySQL connection
try:
    DB = MySQLdb.connect(host=AUTH['host'],
                         user=AUTH['user'],
                         passwd=AUTH['passwd'],
                         db=AUTH['db'])

    # Get the cursor, which is used to traverse the database, line by line
    cursor = DB.cursor()

    print "db connection success"
except:
    print "db connection failed"

# Create the INSERT INTO sql query
# query = """INSERT INTO orders (product, customer_type, rep, date, actual, expected, open_opportunities, closed_opportunities, city, state, zip, population, region) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""

# Iterate through sheet
# start at row 2 bc headers
for r in range(1, sheet.nrows):
    date = sheet.cell(r, 0).value
    TBnum = sheet.cell(r, 1).value
    name = sheet.cell(r, 5).value
    owner = sheet.cell(r, 6).value
    latitude = sheet.cell(r, 7).value
    longitude = sheet.cell(r, 8).value
    description = sheet.cell(r, 9).value
    qty = sheet.cell(r, 10).value

    # Assign values from each row
    values = (latitude, longitude, owner, name, description, date, qty, TBnum)
    print values

#       # Execute sql Query
#       cursor.execute(query, values)

# # Close the cursor
# cursor.close()

# # Commit the transaction
# DB.commit()

# # Close the database connection
# DB.close()

# # Print results
# print ""
# print "All Done! Bye, for now."
# print ""
# columns = str(sheet.ncols)
# rows = str(sheet.nrows)