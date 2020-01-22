import pyodbc
import xlsxwriter
import win32com.client as win32
import os
import sys
from datetime import datetime

def runQuery(date, directory, dsn):
    dirPath = os.getcwd( )
    today = date
    dir = directory

    # Connection Information
    conn = pyodbc.connect(dsn)

    # Create and execute the query
    cursor = conn.cursor( )
    sql = """
        select
        	[MC Surfaces, Inc].[dbo].pchord.odrdby,
        	[MC Surfaces, Inc].[dbo].employ.fulfst,
        	[MC Surfaces, Inc].[dbo].actrec.sprvsr,
        	fieldRep.fulfst,
        	[MC Surfaces, Inc].[dbo].pchord.dscrpt,
        	[MC Surfaces, Inc].[dbo].pchord.orddte,
        	[MC Surfaces, Inc].[dbo].actpay.vndnme,
        	[MC Surfaces, Inc].[dbo].pchord.ordnum,
        	[MC Surfaces, Inc].[dbo].pchord.jobnum,
            [MC Surfaces, Inc].[dbo].actrec.jobnme,
        	[MC Surfaces, Inc].[dbo].pchord.pchttl,
        	[MC Surfaces, Inc].[dbo].pchord.ordtyp,
        	[MC Surfaces, Inc].[dbo].pchord.ntetxt
        	from [MC Surfaces, Inc].[dbo].pchord
        		left join [MC Surfaces, Inc].[dbo].employ
        			on [MC Surfaces, Inc].[dbo].pchord.odrdby = [MC Surfaces, Inc].[dbo].employ.recnum
        		left join [MC Surfaces, Inc].[dbo].actpay
        			on [MC Surfaces, Inc].[dbo].pchord.vndnum = [MC Surfaces, Inc].[dbo].actpay.recnum
        		left join [MC Surfaces, Inc].[dbo].actrec
        			on [MC Surfaces, Inc].[dbo].pchord.jobnum = [MC Surfaces, Inc].[dbo].actrec.recnum
        		left join [MC Surfaces, Inc].[dbo].employ as fieldRep
        			on [MC Surfaces, Inc].[dbo].actrec.sprvsr = fieldRep.recnum
        		where
        			[MC Surfaces, Inc].[dbo].pchord.ordtyp = 2
        			and [MC Surfaces, Inc].[dbo].pchord.orddte <= CAST(DATEADD(DAY, -1, GETDATE( )) as date)
        			and [MC Surfaces, Inc].[dbo].pchord.orddte >= CAST(DATEADD(DAY, -6, (DATEADD(DAY, -1, GETDATE( )))) as date)
                order by [MC Surfaces, Inc].[dbo].pchord.odrdby, [MC Surfaces, Inc].[dbo].actrec.sprvsr, [MC Surfaces, Inc].[dbo].pchord.dscrpt;
    """
    cursor.execute(sql)
    tuples = list(cursor.fetchall( ))
    data = [ ]
    for tuple in tuples:
        data.append(list(tuple))

    # Save the results to .xlsx file
    fileName = "CHARGEABLE REPORT " + today + ".xlsx"
    if os.path.exists(dir + '\\Reports\\' + fileName):
        os.remove(dir + '\\Reports\\' + fileName)
    workbook = xlsxwriter.Workbook(dir + "\\Reports\\" + fileName)
    worksheet = workbook.add_worksheet( )
    worksheet.set_landscape( )
    worksheet.set_header('&C&24Weekly Warranty Report Chargeable')
    worksheet.fit_to_pages(1, 0)

    # Workbook Formatting
    bold = workbook.add_format({'bold': True})
    title = workbook.add_format({'bold': True, 'align': 'right'})
    titleLeft = workbook.add_format({'bold': True, 'align': 'left'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})
    topLeft = workbook.add_format({'valign': 'top', 'align': 'left'})
    money = workbook.add_format({'num_format': '$#,##0.00', 'valign': 'top'})

    # Cell Dimensions
    worksheet.set_column("J:J", 35)
    worksheet.set_row(0, 20)

    # Write information to the workbook
    worksheet.write('A1', 'Ordered By', bold)
    worksheet.write('B1', 'Field Rep', bold)
    worksheet.write('C1', 'Description', bold)
    worksheet.write('D1', 'Order Date', bold)
    worksheet.write('E1', 'Vendor Name', bold)
    worksheet.write('F1', 'Order #', bold)
    worksheet.write('G1', 'Job', titleLeft)
    worksheet.write('H1', 'Total', title)
    worksheet.write('I1', 'Type', bold)
    worksheet.write('J1', 'Notes', bold)

    row = 2
    col = 0
    for e1, ob, e2, fr, dsc, od, vn, on, j, jnm, t, typ, nte in data:
        worksheet.write(row, col, str(e1) + " - " + str(ob), top)
        if e2 is None:
            worksheet.write(row, col+1, "-----", top)
        else:
            worksheet.write(row, col + 1, str(e2) + " - " + str(fr), top)
        worksheet.write(row, col + 2, dsc, top)
        invoice_date = datetime.strptime(od, '%Y-%m-%d')
        date_format = workbook.add_format({'num_format': 'mm/dd/yyyy', 'valign': 'top'})
        worksheet.write(row, col + 3, invoice_date, date_format)
        worksheet.write(row, col + 4, vn, top)
        worksheet.write(row, col + 5, on, top)
        worksheet.write(row, col + 6, str(j) + " - " + str(jnm), topLeft)
        worksheet.write(row, col + 7, t, money)
        worksheet.write(row, col + 8, "Chargeable", top)
        worksheet.write(row, col + 9, nte, wrap)
        row += 1

    # Close the workbook and connection
    workbook.close( )
    conn.close( )

    # Auto format column width
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(dir + "\\Reports\\" + fileName)
    ws = wb.Worksheets("Sheet1")
    ws.Columns.AutoFit( )
    wb.Save( )
    excel.Application.Quit( )
