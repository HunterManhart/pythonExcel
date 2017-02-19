import xlrd
import MySQLdb
import datetime as dt 
from env.auth import AUTH

def conversion(old):
    if is_number(old):
        return float(old)

    direction = {'N':-1, 'S':1, 'E': -1, 'W':1}
    new = old.replace(u'\xb0',' ').replace('\'',' ').replace('"',' ')    
    new = new.split()
    new_dir = new.pop()
    new.extend([0,0,0])
    return (int(new[0])+int(new[1])/60.0+float(new[2])/3600.0) * direction[new_dir]

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def getID(old):
    id = old.split("-")[1]
    return int(id[:5])

def getDate(num):
    dateTup = xlrd.xldate_as_tuple(int(num), 0)
    return dt.date(dateTup[0], dateTup[1], dateTup[2])

# Open the workbook and define the worksheet
book = xlrd.open_workbook("excel/2016.xlsx")
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
query = """INSERT INTO sites (latitude, longitude, owner, site_name, description, ship_date, unit_quantity, boom_type, color, tuffboom_id, user_create, user_update) VALUES (%f, %f, "%s", "%s", "%s", "%s", %d, 1, 1, %d, 2, 2)"""

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

    # Make sure lat and long aren't empty
    if latitude == "" or longitude == "":
        continue

    latitude = conversion(latitude)
    longitude = conversion(longitude)

    TBnum = getID(TBnum)

    date = str(getDate(date))

    # Assign values from each row
    values = (latitude, longitude, owner, name, description, date, int(qty), TBnum)

    print query % values

    # Execute sql Query
    cursor.execute(query % values)

# Close the cursor
cursor.close()

# Commit the transaction
DB.commit()

# Close the database connection
DB.close()
