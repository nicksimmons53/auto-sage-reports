import pyodbc
import datetime as dt
from datetime import datetime
import win32com.client as win32
import os
import sys
import xlsxwriter

def runQuery(directory, dsn):
    dirPath = os.getcwd( )
    today = dt.date.today( ).strftime("%m_%d_%y")
    currentDate = dt.date.today( ).strftime("%d/")
    dir = directory

    # Connection Information
    conn = pyodbc.connect(dsn)

    # Create and execute the query
    cursor = conn.cursor( )
    sql = """
        select
        	[MC Surfaces, Inc].[dbo].actrec.dptmnt,
        	[MC Surfaces, Inc].[dbo].dptmnt.dptnme,
        	[MC Surfaces, Inc].[dbo].reccln.clnnme,
        	[MC Surfaces, Inc].[dbo].acrinv.jobnum,
        	[MC Surfaces, Inc].[dbo].actrec.jobnme,
        	[MC Surfaces, Inc].[dbo].actrec.usrdf2,
        	[MC Surfaces, Inc].[dbo].acrinv.invdte,
        	[MC Surfaces, Inc].[dbo].acrinv.invnum,
        	[MC Surfaces, Inc].[dbo].acrinv.invttl,
        	[MC Surfaces, Inc].[dbo].acrinv.entdte
        	from [MC Surfaces, Inc].[dbo].acrinv
        		left join [MC Surfaces, Inc].[dbo].actrec
        			on [MC Surfaces, Inc].[dbo].acrinv.jobnum = [MC Surfaces, Inc].[dbo].actrec.recnum
        		left join [MC Surfaces, Inc].[dbo].dptmnt
        			on [MC Surfaces, Inc].[dbo].actrec.dptmnt = [MC Surfaces, Inc].[dbo].dptmnt.recnum
        		left join [MC Surfaces, Inc].[dbo].reccln
        			on [MC Surfaces, Inc].[dbo].actrec.clnnum = [MC Surfaces, Inc].[dbo].reccln.recnum
        		where
        			[MC Surfaces, Inc].[dbo].acrinv.status < 5
        			and [MC Surfaces, Inc].[dbo].acrinv.invdte <= CAST(DATEADD(DAY, -1, GETDATE( )) as date)
        			and [MC Surfaces, Inc].[dbo].acrinv.invdte >= CAST(DATEADD(DAY, -6, (DATEADD(DAY, -1, GETDATE( )))) as date)
                order by
                    [MC Surfaces, Inc].[dbo].actrec.dptmnt, [MC Surfaces, Inc].[dbo].reccln.clnnme, [MC Surfaces, Inc].[dbo].acrinv.jobnum;
    """

    cursor.execute(sql)
    tuples = list(cursor.fetchall( ))
    data = [ ]
    for tuple in tuples:
        data.append(list(tuple))

    # Save the results to .xlsx file
    fileName = "WEEKLY REVENUE REPORT " + today + ".xlsx"
    if os.path.exists(dir + '\\Reports\\' + fileName):
        os.remove(dir + '\\Reports\\' + fileName)
    workbook = xlsxwriter.Workbook(dir + "\\Reports\\" + fileName)
    worksheet = workbook.add_worksheet( )
    worksheet.set_landscape( )
    worksheet.set_header('&C&24Weekly A/R Invoices Report')
    worksheet.fit_to_pages(1, 0)

    # Workbook Formatting
    bold = workbook.add_format({'bold': True})
    title = workbook.add_format({'bold': True, 'align': 'right'})
    titleLeft = workbook.add_format({'bold': True, 'align': 'Left'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})
    money = workbook.add_format({'num_format': '#,##0.00', 'align': 'right'})
    moneyBold = workbook.add_format({'num_format': '#,##0.00', 'align': 'right', 'bold': True})

    # Cell Dimensions
    worksheet.set_column("J:J", 35)
    worksheet.set_row(0, 20)

    # Write information to the workbook
    worksheet.write('A1', 'Department', bold)
    worksheet.write('B1', 'Client Name', bold)
    worksheet.write('C1', 'Job #', title)
    worksheet.write('D1', 'Job Name', bold)
    worksheet.write('E1', 'Neighborhood', bold)
    worksheet.write('F1', 'Invoice Date', titleLeft)
    worksheet.write('G1', 'Invoice #', title)
    worksheet.write('H1', '', title)
    worksheet.write('I1', 'Invoice Total', title)
    worksheet.write('J1', 'Entered', title)

    row = 2
    col = 0
    total = 0
    for d, dn, cn, j, jnm, nb, id, i, it, e in data:
        worksheet.write(row, col, str(d) + " - " + dn, top)
        worksheet.write(row, col + 1, cn, top)
        worksheet.write(row, col + 2, j, top)
        worksheet.write(row, col + 3, jnm, top)
        worksheet.write(row, col + 4, nb, top)
        invoice_date = datetime.strptime(id, '%Y-%m-%d')
        date_format = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left'})
        worksheet.write(row, col + 5, invoice_date, date_format)
        worksheet.write(row, col + 6, i, top)
        worksheet.write(row, col + 7, '$', top)
        if float(it) < 0:
            worksheet.write(row, col + 8, '(' + str(it) + ')', money)
            total += it
        else:
            worksheet.write(row, col + 8, it, money)
            total += it

        entered_date = datetime.strptime(e, '%Y-%m-%d')
        worksheet.write(row, col + 9, entered_date, date_format)
        row += 1

    worksheet.write(row + 1, col + 7, '$', bold)
    worksheet.write(row + 1, col + 8, total, moneyBold)

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
