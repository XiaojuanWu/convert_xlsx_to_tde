import csv
import time
import os
import datetime
import locale
import re
import configparser
import pandas as pd
import argparse
from tableausdk import Extract as tde
from tableausdk import *
from tableau_rest_api.tableau_rest_api import *

# Define type maps
schemaIniTypeMap = {
    'Bit': Type.BOOLEAN,
    'Byte': Type.INTEGER,
    'Short': Type.INTEGER,
    'Long': Type.INTEGER,
    'Integer': Type.INTEGER,
    'Single': Type.DOUBLE,
    'Double': Type.DOUBLE,
    'Date': Type.DATE,
    'DateTime': Type.DATETIME,
    'Text': Type.UNICODE_STRING,
    'Memo': Type.UNICODE_STRING
}

fieldSetterMap = {
    Type.BOOLEAN: lambda row, colNo, value: row.setBoolean(colNo, value.lower() == "true"),
    Type.INTEGER: lambda row, colNo, value: row.setInteger(colNo, int(value)),
    Type.DOUBLE: lambda row, colNo, value: row.setDouble(colNo, float(value)),
    Type.UNICODE_STRING: lambda row, colNo, value: row.setString(colNo, value),
    Type.CHAR_STRING: lambda row, colNo, value: row.setCharString(colNo, value),
    Type.DATE: lambda row, colNo, value: conversion.setDate(row, colNo, value),
    Type.DATETIME: lambda row, colNo, value: conversion.setDateTime(row, colNo, value)
}


class generate_tde(object):

    def __init__(self, xlsx_file):
        self.xlsx_file = xlsx_file

    def clear_old_file(self):
        if os.path.exists('results.tde'):
            print '### Deleting old results.tde file'
            os.remove('results.tde')

    def xls2csv(self, xlsx_file):
        # Read file specified in prompt
        print '### Reading Excel File'
        data_xls = pd.ExcelFile(xlsx_file)
        sheets = data_xls.sheet_names
        print '### Sheets found:', sheets
        df = pd.DataFrame()
        #bring all of the data into a single pandas data frame
        for sheet in sheets:
            df = df.append(pd.read_excel(xlsx_file, sheet))
        print '### Generating results.csv'
        print '### Columns found:', list(df.columns.values)
        # Rearrange columns in data frame
        cols = list(df.columns.values)
        cols = cols[3:4] + cols[:1] + cols[-1:] + cols[1:2] + cols[2:3]
        print '### New order:', cols
        # Move the data frame into a csv
        df.to_csv('out.csv')

        print '### Removing unwanted characters'

        def replace_all(self, text, dic):
            for x, y in dic.iteritems():
                text = text.replace(x, y)
            return text
        # Clean up the data in the csv
        with open('out.csv', 'r') as f:
            a_dic = {'$-   ': '0'}
            text = f.read()
            with open('out2.csv', 'w') as w:
                text = replace_all(self, text, a_dic)
                w.write(text)

        print '### Removing extra column'
        # Finalize column order and position
        with open("out2.csv", "rb") as source:
            rdr = csv.reader(source)
            with open("results.csv", "wb") as result:
                wtr = csv.writer(result)
                for r in rdr:
                    wtr.writerow((r[4], r[1], r[5], r[2], r[3]))
        # Delete temp files
        os.remove("out.csv")
        os.remove("out2.csv")
        print '### Csv cleanup complete'

    def setDate(self, row, colNo, value):
        d = datetime.datetime.strptime(value, "%Y-%m-%d")
        row.setDate(colNo, d.year, d.month, d.day)

    def setDateTime(self, row, colNo, value):
        if(value.find(".") != -1):
            d = datetime.datetime.strptime(value, "%Y-%m-%d %H:%M:%S.%f")
        else:
            d = datetime.datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
        row.setDateTime(colNo, d.year, d.month, d.day, d.hour, d.minute, d.second, d.microsecond / 100)

    def create_tde(self):
        # Identify CSV input
        csvFile = 'results.csv'
        # Open CSV file
        csvReader = csv.reader(open(csvFile, 'rb'), delimiter=',', quotechar='"')
        # Read schema.ini file, if it exists
        schemaFile = "schema.ini"
        hasHeader = True
        colNames = []
        colTypes = []
        locale.setlocale(locale.LC_ALL, '')
        colParser = re.compile(r'(col)(\d+)', re.IGNORECASE)
        schemaIni = configparser.ConfigParser()
        schemaIni.read(schemaFile)
        schemaIni.sections()
        for item in schemaIni.items('default'):
            name = item[0]
            value = item[1]
            if name == "colnameheader":
                hasHeader = value == "True"
            m = colParser.match(name)
            if not m:
                continue
            colName = m.groups()[0]
            colNo = int(m.groups()[1]) - 1
            parts = value.split(' ')
            name = parts[0]
            try:
                type = schemaIniTypeMap[parts[1]]
            except KeyError:
                type = Type.UNICODE_STRING
            while colNo >= len(colNames):
                colNames.append(None)
                colTypes.append(Type.UNICODE_STRING)
            colNames[colNo] = name
            colTypes[colNo] = type

        # Create TDE output
        tdefile = csvFile.split('.')[0] + ".tde"
        print "### Creating extract:", tdefile
        with tde.Extract(tdefile) as extract:
            table = None  # set by createTable
            tableDef = None

            # Define createTable function
            def createTable(line):
                if line:
                    # append with empty columns so we have the same number of columns as the header row
                    while len(colNames) < len(line):
                        colNames.append(None)
                        colTypes.append(Type.UNICODE_STRING)
                    # write in the column names from the header row
                    colNo = 0
                    for colName in line:
                        colNames[colNo] = colName
                        colNo += 1

                # for any unnamed column, provide a default
                for i in range(0, len(colNames)):
                    if colNames[i] is None:
                        colNames[i] = 'F' + str(i + 1)

                # create the schema and the table from it
                if extract.hasTable('Extract'):
                    table = extract.openTable('Extract')
                    tableDef = table.getTableDefinition()
                else:
                    tableDef = tde.TableDefinition()
                    for i in range(0, len(colNames)):
                        tableDef.addColumn(colNames[i], colTypes[i])
                    table = extract.addTable("Extract", tableDef)
                return table, tableDef

            # Read the table
            print "### Adding rows to .tde..."
            rowNo = 0
            for line in csvReader:
                # Create the table upon first row (which may be a header)
                if table is None:
                    table, tableDef = createTable(line if hasHeader else None)
                    if hasHeader:
                        continue

                # We have a table, now write a row of values
                row = tde.Row(tableDef)
                colNo = 0
                for field in line:
                    if(colTypes[colNo] != Type.UNICODE_STRING and field == ""):
                        row.setNull(colNo)
                    else:
                        fieldSetterMap[colTypes[colNo]](row, colNo, field)
                    colNo += 1
                table.insert(row)

            print "### All rows added"

    def publish_tde(xlsx_file, xlsx_name, site_name):
        # REST api upload process
        server = 'http://p-p3tableaum01.use01.plat.priv'
        username = '******'
        password = '******'
        rest = TableauRestApi(server, username, password, site_content_url=site_name)
        logger = Logger('Publish.log')
        rest.enable_logging(logger)
        rest.signin()

        print "### Publishing to server..."
        template_ds = 'results.tde'
        filename = xlsx_name
        print '### Removing temp files'

        rest_ds_proj_luid = rest.query_project_luid_by_name('default')
        # Publishing data source to site, from disk
        rest.publish_datasource(template_ds, filename + '.tde', rest_ds_proj_luid, overwrite=True, connection_password=password, connection_username=username)
        print "### Successfully uploaded", filename + '.tde'


if __name__ == '__main__':
    # Create variables to fill via command line
    parser = argparse.ArgumentParser(description='Converts .xlsx to .tde')
    parser.add_argument('-f', action='store', dest='excelfilename', help='Name of the file to be processed')
    parser.add_argument('-s', action='store', dest='sitename', help='Site where the file is to be uploaded')
    parser_results = parser.parse_args()

    xlsx_file = parser_results.excelfilename
    # Truncate file extension
    if xlsx_file.endswith('.xlsx'):
        xlsx_name = xlsx_file[:-5]
    else:
        xlsx_name = xlsx_file
    site_name = parser_results.sitename

    startTime = time.time()
    # Call the class
    conversion = generate_tde(xlsx_file)
    # Start the process
    conversion.clear_old_file()
    # Step 1
    conversion.xls2csv(xlsx_file)
    q_endTime = time.time()
    q_runTime = q_endTime - startTime
    e_startTime = time.time()
    # Step 2
    conversion.create_tde()
    e_runTime = time.time() - e_startTime
    u_startTime = time.time()
    # Step 3
    conversion.publish_tde(xlsx_name, site_name)
    u_runTime = time.time() - u_startTime
    # Count the rows
    readoutfile = open('results.csv', 'r')
    rowcount = sum(1 for row in readoutfile) - 1
    readoutfile.close()
    os.remove('results.csv')

# Output elapsed time
    print "----------------------------------------------------------------------"
    print xlsx_name, "Elapsed Time Summary:"
    print "Step 1: Read xlsx\t", locale.format("%.2f", q_runTime), "seconds"
    print "Step 2: Generate .tde\t", locale.format("%.2f", e_runTime), "seconds"
    print "Step 3: Upload\t\t", locale.format("%.2f", u_runTime), "seconds"
    print "----------------------------------------------------------------------"
    print "Elapsed:\t\t", locale.format("%.2f", time.time() - startTime), "seconds"
    print "Rows Returned:\t\t", rowcount
    print "----------------------------------------------------------------------"
    print "----------------------------------------------------------------------"
