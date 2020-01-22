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
        	[MC Surfaces, Inc].[dbo].srvsch.vndnum,
        	[MC Surfaces, Inc].[dbo].actpay.vndnme,
        	[MC Surfaces, Inc].[dbo].srvinv.recnum,
        	[MC Surfaces, Inc].[dbo].srvinv.schdte,
        	[MC Surfaces, Inc].[dbo].srvinv.entdte,
        	[MC Surfaces, Inc].[dbo].srvinv.invtyp,
        	[MC Surfaces, Inc].[dbo].srvtyp.typnme,
        	[MC Surfaces, Inc].[dbo].reccln.clnnme,
        	[MC Surfaces, Inc].[dbo].srvinv.jobnum,
        	[MC Surfaces, Inc].[dbo].actrec.jobnme,
        	[MC Surfaces, Inc].[dbo].srvinv.dscrpt,
        	[MC Surfaces, Inc].[dbo].srvinv.ntetxt,
        	[MC Surfaces, Inc].[dbo].srvinv.usrnme,
        	[MC Surfaces, Inc].[dbo].actrec.dptmnt
        	from [MC Surfaces, Inc].[dbo].srvinv
        		left join [MC Surfaces, Inc].[dbo].reccln
        			on [MC Surfaces, Inc].[dbo].srvinv.clnnum = [MC Surfaces, Inc].[dbo].reccln.recnum
        		left join [MC Surfaces, Inc].[dbo].srvtyp
        			on [MC Surfaces, Inc].[dbo].srvinv.invtyp = [MC Surfaces, Inc].[dbo].srvtyp.recnum
        		left join [MC Surfaces, Inc].[dbo].srvsch
        			on [MC Surfaces, Inc].[dbo].srvinv.recnum = [MC Surfaces, Inc].[dbo].srvsch.recnum
        		left join [MC Surfaces, Inc].[dbo].actpay
        			on [MC Surfaces, Inc].[dbo].srvsch.vndnum = [MC Surfaces, Inc].[dbo].actpay.recnum
        		left join [MC Surfaces, Inc].[dbo].actrec
        			on [MC Surfaces, Inc].[dbo].srvinv.jobnum = [MC Surfaces, Inc].[dbo].actrec.recnum
        		where
        			[MC Surfaces, Inc].[dbo].actrec.dptmnt = 200
        			and [MC Surfaces, Inc].[dbo].srvinv.invtyp != 2
        			and [MC Surfaces, Inc].[dbo].srvsch.schdte <= CAST(DATEADD(DAY, -1, GETDATE( )) as date)
        			and [MC Surfaces, Inc].[dbo].srvsch.schdte >= CAST(DATEADD(DAY, -6, (DATEADD(DAY, -1, GETDATE( )))) as date)
        		order by [MC Surfaces, Inc].[dbo].srvinv.dscrpt, [MC Surfaces, Inc].[dbo].actpay.vndnme, [MC Surfaces, Inc].[dbo].srvinv.schdte;
    """

    cursor.execute(sql)
    tuples = cursor.fetchall( )
    data = [ ]
    for tuple in list(tuples):
        data.append(list(tuple))

    # Save the results to .xlsx file
    fileName = "HOUSTON REPAIRS BY WARRANTY TECH " + today + ".xlsx"
    if os.path.exists(dir + '\\Reports\\' + fileName):
        os.remove(dir + '\\Reports\\' + fileName)
    workbook = xlsxwriter.Workbook(dir + "\\Reports\\" + fileName)
    worksheet = workbook.add_worksheet( )
    worksheet.set_landscape( )
    worksheet.set_header('&C&24Weekly Warranty Report by Tech')
    worksheet.fit_to_pages(1, 0)

    # Workbook Formatting
    bold = workbook.add_format({'bold': True})
    title = workbook.add_format({'bold': True, 'align': 'right'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})
    titleLeft = workbook.add_format({'bold': True, 'align': 'Left'})

    # Cell Dimensions
    worksheet.set_column("I:I", 35)
    worksheet.set_column("H:H", 35)
    worksheet.set_row(0, 20)

    # Write information to the workbook
    worksheet.write('A1', 'Vendor Name', bold)
    worksheet.write('B1', 'Record #', bold)
    worksheet.write('C1', 'Sched. Date', bold)
    worksheet.write('D1', 'Entered', bold)
    worksheet.write('E1', 'Type', bold)
    worksheet.write('F1', 'Client Name', bold)
    worksheet.write('G1', 'Job', titleLeft)
    worksheet.write('H1', 'Description', titleLeft)
    worksheet.write('I1', 'Notes', bold)
    worksheet.write('J1', 'Dept.', bold)

    row = 2
    col = 0
    occupied_houses = [ ]
    info = [ ]
    for vn, vnm, rn, sd, ed, it, tn, cn, jn, jnm, dsc, nt, us, dpt in data:
        info = [vn, vnm, rn, sd, ed, it, tn, cn, jn, jnm, dsc, nt, us, dpt]
        if "CASA OCUPADA" in nt:
            occupied_houses.append(info)

        if vn is None:
            worksheet.write(row, col, "-----", top)
        else:
            worksheet.write(row, col, vnm, top)
        worksheet.write(row, col + 1, rn, top)
        sched_date = datetime.strptime(sd, '%Y-%m-%d')
        date_format = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left', 'valign': 'top'})
        worksheet.write(row, col + 2, sched_date, date_format)
        enter_date = datetime.strptime(ed, '%Y-%m-%d')
        worksheet.write(row, col + 3, enter_date, date_format)
        worksheet.write(row, col + 4, str(it) + " - " + tn, top)
        worksheet.write(row, col + 5, cn, top)
        worksheet.write(row, col + 6, str(jn) + " - " + jnm, top)
        worksheet.write(row, col + 7, dsc, top)
        worksheet.write(row, col + 8, nt, wrap)
        worksheet.write(row, col + 9, str(dpt) + " - Houston", top)
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



    ############ OCCUPIED HOUSES ############
    # Save the results to .xlsx file
    fileName = "OCCUPIED HOME REPAIRS " + today + ".xlsx"
    if os.path.exists(dir + '\\Reports\\' + fileName):
        os.remove(dir + '\\Reports\\' + fileName)
    workbook = xlsxwriter.Workbook(dir + "\\Reports\\" + fileName)
    worksheet = workbook.add_worksheet( )
    worksheet.set_landscape( )
    worksheet.set_header('&C&24Weekly Warranty Report by Tech - Occupied Homes')
    worksheet.fit_to_pages(1, 0)

    # Workbook Formatting
    bold = workbook.add_format({'bold': True})
    title = workbook.add_format({'bold': True, 'align': 'right'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})
    titleLeft = workbook.add_format({'bold': True, 'align': 'Left'})
    highlight = workbook.add_format({'bold': True})
    highlight.set_pattern(1)
    highlight.set_bg_color('yellow')

    # Cell Dimensions
    worksheet.set_column("I:I", 35)
    worksheet.set_column("H:H", 35)
    worksheet.set_row(0, 20)

    # Write information to the workbook
    worksheet.write('A1', 'Vendor Name', bold)
    worksheet.write('B1', 'Record #', bold)
    worksheet.write('C1', 'Sched. Date', bold)
    worksheet.write('D1', 'Entered', bold)
    worksheet.write('E1', 'Type', bold)
    worksheet.write('F1', 'Client Name', bold)
    worksheet.write('G1', 'Job', titleLeft)
    worksheet.write('H1', 'Description', titleLeft)
    worksheet.write('I1', 'Notes', bold)
    worksheet.write('J1', 'Dept.', bold)

    row = 2
    col = 0
    count = 0
    for vn, vnm, rn, sd, ed, it, tn, cn, jn, jnm, dsc, nt, us, dpt in occupied_houses:
        if vn is None:
            worksheet.write(row, col, "-----", top)
        else:
            worksheet.write(row, col, vnm, top)
        worksheet.write(row, col + 1, rn, top)
        sched_date = datetime.strptime(sd, '%Y-%m-%d')
        date_format = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'left', 'valign': 'top'})
        worksheet.write(row, col + 2, sched_date, date_format)
        enter_date = datetime.strptime(ed, '%Y-%m-%d')
        worksheet.write(row, col + 3, enter_date, date_format)
        worksheet.write(row, col + 4, str(it) + " - " + tn, top)
        worksheet.write(row, col + 5, cn, top)
        worksheet.write(row, col + 6, str(jn) + " - " + jnm, top)
        worksheet.write(row, col + 7, dsc, top)
        worksheet.write(row, col + 8, nt, wrap)
        worksheet.write(row, col + 9, str(dpt) + " - Houston", top)
        count += 1
        row += 1

    worksheet.write(row + 1, col + 7, str(count) + " Occupied Repairs", highlight)

    # Close the workbook and connection
    workbook.close( )

    # Auto format column width
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(dir + "\\Reports\\" + fileName)
    ws = wb.Worksheets("Sheet1")
    ws.Columns.AutoFit( )
    wb.Save( )
    excel.Application.Quit( )
