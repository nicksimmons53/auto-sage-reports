import pyodbc
import xlsxwriter
import win32com.client as win32
import os
import sys

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
        	[MC Surfaces, Inc].[dbo].acrinv.invnum,
        	[MC Surfaces, Inc].[dbo].acrinv.invdte,
        	[MC Surfaces, Inc].[dbo].acrinv.jobnum,
        	[MC Surfaces, Inc].[dbo].actrec.jobnme,
        	[MC Surfaces, Inc].[dbo].actrec.dptmnt,
        	[MC Surfaces, Inc].[dbo].dptmnt.dptnme,
        	[MC Surfaces, Inc].[dbo].acrinv.dscrpt,
        	[MC Surfaces, Inc].[dbo].acrinv.usrdf2,
        	[MC Surfaces, Inc].[dbo].acrinv.invtyp,
        	[MC Surfaces, Inc].[dbo].acrinv.status,
        	[MC Surfaces, Inc].[dbo].acrinv.invttl
        	from [MC Surfaces, Inc].[dbo].acrinv
        		left join [MC Surfaces, Inc].[dbo].actrec
        			on [MC Surfaces, Inc].[dbo].acrinv.jobnum = [MC Surfaces, Inc].[dbo].actrec.recnum
        		left join [MC Surfaces, Inc].[dbo].dptmnt
        			on [MC Surfaces, Inc].[dbo].actrec.dptmnt = [MC Surfaces, Inc].[dbo].dptmnt.recnum
        		where
        			[MC Surfaces, Inc].[dbo].acrinv.dscrpt = 'TEMPLATE ERROR- LOSS'
        				and [MC Surfaces, Inc].[dbo].acrinv.invdte <= CAST(DATEADD(DAY, -1, GETDATE( )) as date)
        				and [MC Surfaces, Inc].[dbo].acrinv.invdte >= CAST(DATEADD(DAY, -6, (DATEADD(DAY, -1, GETDATE( )))) as date)
        			or [MC Surfaces, Inc].[dbo].acrinv.dscrpt = '2ND TIME TEMPLATE ERROR- LOSS'
        				and [MC Surfaces, Inc].[dbo].acrinv.invdte <= CAST(DATEADD(DAY, -1, GETDATE( )) as date)
        				and [MC Surfaces, Inc].[dbo].acrinv.invdte >= CAST(DATEADD(DAY, -6, (DATEADD(DAY, -1, GETDATE( )))) as date);
    """
    cursor.execute(sql)
    tuples = list(cursor.fetchall( ))
    data = [ ]
    for tuple in tuples:
        data.append(list(tuple))

    # Save the results to .xlsx file
    fileName = "TEMPLATE LOSS REPORT " + today + ".xlsx"
    if os.path.exists(dir + '\\Reports\\' + fileName):
        os.remove(dir + '\\Reports\\' + fileName)
    workbook = xlsxwriter.Workbook(dir + "\\Reports\\" + fileName)
    worksheet = workbook.add_worksheet( )
    worksheet.set_landscape( )
    worksheet.set_header('&C&24Weekly Template Loss Report')
    worksheet.fit_to_pages(1, 0)

    # Workbook Formatting
    bold = workbook.add_format({'bold': True})
    title = workbook.add_format({'bold': True, 'align': 'right'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})
    money = workbook.add_format({'num_format': '$#,##0.00'})

    # Cell Dimensions
    worksheet.set_column("J:J", 35)
    worksheet.set_row(0, 20)

    # Write information to the workbook
    worksheet.write('A1', 'Invoice #', bold)
    worksheet.write('B1', 'Invoice Date', bold)
    worksheet.write('C1', 'Job', bold)
    worksheet.write('D1', 'Department', bold)
    worksheet.write('E1', 'Description', bold)
    worksheet.write('F1', 'User Defined', bold)
    worksheet.write('G1', 'Type', title)
    worksheet.write('H1', 'Status', title)
    worksheet.write('I1', 'Invoice Total', title)

    row = 2
    col = 0
    for i, id, jn, jnm, d, dpt, dsc, ud, it, st, ttl in data:
        worksheet.write(row, col, i, top)
        worksheet.write(row, col + 1, id, top)
        worksheet.write(row, col + 2, str(jn) + " - " + jnm, top)
        worksheet.write(row, col + 3, str(d) + " - " + dpt, top)
        worksheet.write(row, col + 4, dsc, top)
        worksheet.write(row, col + 5, ud, top)
        worksheet.write(row, col + 6, it, top)
        worksheet.write(row, col + 7, st, top)
        worksheet.write(row, col + 8, ttl, money)
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
