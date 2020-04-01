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
        	[MCS INC].[dbo].tkflin.prtnum,
        	[MCS INC].[dbo].tkfprt.prtnme,
        	[MCS INC].[dbo].tkflin.extqty
        	from [MCS INC].[dbo].tkflin
        		left join [MCS INC].[dbo].ptotkf
        			on [MCS INC].[dbo].tkflin.recnum = [MCS INC].[dbo].ptotkf.recnum
        		left join [MCS INC].[dbo].actrec
        			on [MCS INC].[dbo].ptotkf.recnum = [MCS INC].[dbo].actrec.recnum
        		left join [MCS INC].[dbo].reccln
        			on [MCS INC].[dbo].actrec.clnnum = [MCS INC].[dbo].reccln.recnum
        		left join [MCS INC].[dbo].tkfprt
        			on [MCS INC].[dbo].tkflin.prtnum = [MCS INC].[dbo].tkfprt.recnum
        		left join [MCS INC].[dbo].prtcls
        			on [MCS INC].[dbo].tkfprt.prtcls = [MCS INC].[dbo].prtcls.recnum
        		where
        			[MCS INC].[dbo].prtcls.recnum = 1100
        			and [MCS INC].[dbo].actrec.biddte <= CAST(DATEADD(DAY, -1, GETDATE( )) as date)
        			and [MCS INC].[dbo].actrec.biddte >= CAST(DATEADD(DAY, -6, (DATEADD(DAY, -1, GETDATE( )))) as date)
        	order by [MCS INC].[dbo].tkfprt.recnum, [MCS INC].[dbo].ptotkf.recnum;
    """
    cursor.execute(sql)
    tuples = list(cursor.fetchall( ))
    data = [ ]
    for tuple in tuples:
        data.append(list(tuple))

    # Save the results to .xlsx file
    fileName = "WEEKLY GRANITE REPORT " + today + ".xlsx"
    if os.path.exists(dir + '\\Reports\\' + fileName):
        os.remove(dir + '\\Reports\\' + fileName)
    workbook = xlsxwriter.Workbook(dir + "\\Reports\\" + fileName)
    worksheet = workbook.add_worksheet( )
    worksheet.set_header('&C&24Weekly Granite Report')
    worksheet.fit_to_pages(1, 0)

    # Workbook Formatting
    bold = workbook.add_format({'bold': True})
    title = workbook.add_format({'bold': True, 'align': 'right'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})
    decimal = workbook.add_format({'num_format': '0.0000', 'bold': True})
    percent_format = workbook.add_format({'num_format': '0.00"%"', 'bold': True, 'align': 'right'})

    # Cell Dimensions
    worksheet.set_column("J:J", 35)
    worksheet.set_row(0, 20)

    # Write information to the workbook
    worksheet.write('A1', 'Part #', bold)
    worksheet.write('B1', 'Ext Quantity', bold)
    worksheet.write('C1', '%', title)

    row = 2
    col = 0
    num = 0
    total = 0.00
    sum = 0.00
    info = [ ]
    partNumber = 0
    partName = ''
    for part_num, part_name, quantity in data:
        if num == 0:
            num = part_num
        elif num != part_num:
            info.append([part_num, part_name, total])
            row += 1
            num = part_num
            total = 0

        total += float(quantity)
        sum += float(quantity)
        partNumber = part_num
        partName = part_name
    info[-1][2] += total

    percentage_sum = 0.00
    row = 2
    sorted_info = sorted(info, key = lambda i : i[2], reverse=True)
    for i in sorted_info:
        percentage = (float(i[2])/sum) * 100
        worksheet.write(row, col, str(i[0]) + " " + i[1], bold)
        worksheet.write(row, col + 1, i[2], decimal)
        worksheet.write(row, col + 2, percentage, percent_format)
        percentage_sum += percentage
        row += 1

    worksheet.write(row + 1, col, 'Grand Total', bold)
    worksheet.write(row + 1, col + 1, sum, bold)
    worksheet.write(row + 1, col + 2, percentage_sum, percent_format)

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
