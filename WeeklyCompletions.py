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
        	[MCS INC].[dbo].pchord.odrdby,
        	[MCS INC].[dbo].employ.fulfst,
        	[MCS INC].[dbo].actrec.sprvsr,
        	fieldRep.fulfst,
        	[MCS INC].[dbo].pchord.dscrpt,
        	[MCS INC].[dbo].pchord.orddte,
        	[MCS INC].[dbo].actpay.vndnme,
        	[MCS INC].[dbo].pchord.ordnum,
        	[MCS INC].[dbo].pchord.jobnum,
        	[MCS INC].[dbo].pchord.pchttl,
        	[MCS INC].[dbo].pchord.ordtyp,
        	[MCS INC].[dbo].pchord.ntetxt
        	from [MCS INC].[dbo].pchord
        		left join [MCS INC].[dbo].employ
        			on [MCS INC].[dbo].pchord.odrdby = [MCS INC].[dbo].employ.recnum
        		left join [MCS INC].[dbo].actpay
        			on [MCS INC].[dbo].pchord.vndnum = [MCS INC].[dbo].actpay.recnum
        		left join [MCS INC].[dbo].actrec
        			on [MCS INC].[dbo].pchord.jobnum = [MCS INC].[dbo].actrec.recnum
        		left join [MCS INC].[dbo].employ as fieldRep
        			on [MCS INC].[dbo].actrec.sprvsr = fieldRep.recnum
        		where
        			[MCS INC].[dbo].pchord.ordtyp = 7
        			and orddte <= CAST(DATEADD(DAY, -1, GETDATE( )) as date)
        			and orddte >= CAST(DATEADD(DAY, -6, (DATEADD(DAY, -1, GETDATE( )))) as date);
    """
    cursor.execute(sql)
    tuples = list(cursor.fetchall( ))
    data = [ ]
    for tuple in tuples:
        data.append(list(tuple))

    # Save the results to .xlsx file
    fileName = "COMPLETION REPORT " + today + ".xlsx"
    if os.path.exists(dir + '\\Reports\\' + fileName):
        os.remove(dir + '\\Reports\\' + fileName)
    workbook = xlsxwriter.Workbook(dir + "\\Reports\\" + fileName)
    worksheet = workbook.add_worksheet( )
    worksheet.set_landscape( )
    worksheet.set_header('&C&24Weekly Completion Report')
    worksheet.fit_to_pages(1, 0)

    # Workbook Formatting
    bold = workbook.add_format({'bold': True})
    title = workbook.add_format({'bold': True, 'align': 'right'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})

    # Cell Dimensions
    worksheet.set_column("J:J", 35)
    worksheet.set_row(0, 20)

    # Write information to the workbook
    worksheet.write('A1', 'Ordered By', bold)
    worksheet.write('B1', 'Field Rep', bold)
    worksheet.write('C1', 'Description', bold)
    worksheet.write('D1', 'Order Date', bold)
    worksheet.write('E1', 'Vendor Name', bold)
    worksheet.write('F1', 'Order Number', bold)
    worksheet.write('G1', 'Job', title)
    worksheet.write('H1', 'Total', title)
    worksheet.write('I1', 'Type', bold)
    worksheet.write('J1', 'Notes', bold)

    row = 2
    col = 0
    for e1, ob, e2, fr, dsc, od, vn, on, j, t, typ, nte in data:
        worksheet.write(row, col, str(e1) + " - " + str(ob), top)
        if e2 is None:
            worksheet.write(row, col+1, "-----", top)
        else:
            worksheet.write(row, col + 1, str(e2) + " - " + str(fr), top)
        worksheet.write(row, col + 2, dsc, top)
        worksheet.write(row, col + 3, od, top)
        worksheet.write(row, col + 4, vn, top)
        worksheet.write(row, col + 5, on, top)
        worksheet.write(row, col + 6, j, top)
        worksheet.write(row, col + 7, t, top)
        worksheet.write(row, col + 8, "Expeditor Completion", top)
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
