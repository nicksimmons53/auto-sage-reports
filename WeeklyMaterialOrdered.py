import pyodbc
import xlsxwriter
import win32com.client as win32
import os
import sys
from datetime import datetime

def getDept(value, user):
    if value == 1:
        if user == "Karinag":
            return "T"
        if user == "GraniteWH":
            return "G"
        if user == "Florenciaf":
            return "T"
        if user == "Katw":
            return "T"
        if user == "Kimc":
            return "T"
        if user == "Orfilap":
            return "T"
        if user == "Edythc":
            return "G"
        if user == "Heatherw":
            return "G"
        if user == "Hollyw":
            return "G"
        if user == "Natalieh":
            return "G"
        if user == "Shandyt":
            return "G"
        if user == "Evelynm":
            return "T"
        if user == "Priscilap":
            return "T"
    elif value == 2:
        if user == "Karinag":
            return "T"
        if user == "GraniteWH":
            return "G"
        if user == "Florenciaf":
            return "T"
        if user == "Katw":
            return "T"
        if user == "Kimc":
            return "T"
        if user == "Orfilap":
            return "T"
        if user == "Edythc":
            return "G"
        if user == "Heatherw":
            return "G"
        if user == "Hollyw":
            return "G"
        if user == "Natalieh":
            return "G"
        if user == "Shandyt":
            return "G"
        if user == "Evelynm":
            return "T"
        if user == "Priscilap":
            return "T"
    elif value >= 100 and value < 1000:
        return "T"
    elif value >= 1000 and value < 3000:
        return "G"
    elif value >= 3000 and value < 4000:
        return "G"
    elif value >= 4000 and value < 5000:
        return "W"
    elif value >= 5000 and value < 5500:
        return "G"
    elif value >= 5500 and value < 6000:
        return "T"
    elif value >= 6000 and value < 7000:
        return "M"
    elif value >= 7000 and value < 8000:
        return "C"
    elif value >= 8000 and value < 9000:
        return "P"
    elif value >= 9000:
        return "OS"

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
        	[MC Surfaces, Inc].[dbo].pchord.ordnum,
        	[MC Surfaces, Inc].[dbo].pchord.vndnum,
        	[MC Surfaces, Inc].[dbo].actpay.vndnme,
        	[MC Surfaces, Inc].[dbo].tkfprt.prtcls,
        	[MC Surfaces, Inc].[dbo].prtcls.clsnme,
        	[MC Surfaces, Inc].[dbo].pcorln.prtnum,
        	[MC Surfaces, Inc].[dbo].pcorln.prtdsc,
        	[MC Surfaces, Inc].[dbo].pcorln.linqty,
        	[MC Surfaces, Inc].[dbo].pcorln.extttl,
        	[MC Surfaces, Inc].[dbo].pchord.entdte,
        	[MC Surfaces, Inc].[dbo].pchord.orddte,
        	[MC Surfaces, Inc].[dbo].pchord.usrnme,
            [MC Surfaces, Inc].[dbo].prtcls.parcls
        	from [MC Surfaces, Inc].[dbo].pchord
        		left join [MC Surfaces, Inc].[dbo].pcorln
        			on [MC Surfaces, Inc].[dbo].pchord.recnum = [MC Surfaces, Inc].[dbo].pcorln.recnum
        		left join [MC Surfaces, Inc].[dbo].tkfprt
        			on [MC Surfaces, Inc].[dbo].pcorln.prtnum = [MC Surfaces, Inc].[dbo].tkfprt.recnum
        		left join [MC Surfaces, Inc].[dbo].actpay
        			on [MC Surfaces, Inc].[dbo].pchord.vndnum = [MC Surfaces, Inc].[dbo].actpay.recnum
        		left join [MC Surfaces, Inc].[dbo].prtcls
        			on [MC Surfaces, Inc].[dbo].tkfprt.prtcls = [MC Surfaces, Inc].[dbo].prtcls.recnum
        		where
        			[MC Surfaces, Inc].[dbo].pchord.vndnum != 1164
                    and [MC Surfaces, Inc].[dbo].actpay.vndtyp != 201
        			and [MC Surfaces, Inc].[dbo].pchord.status <= 4
        			and [MC Surfaces, Inc].[dbo].pchord.orddte <= CAST(DATEADD(DAY, -1, GETDATE( )) as date)
        			and [MC Surfaces, Inc].[dbo].pchord.orddte >= CAST(DATEADD(DAY, -6, (DATEADD(DAY, -1, GETDATE( )))) as date)
        			and [MC Surfaces, Inc].[dbo].pchord.entdte <= CAST(DATEADD(DAY, -1, GETDATE( )) as date)
        			and [MC Surfaces, Inc].[dbo].pchord.entdte >= CAST(DATEADD(DAY, -6, (DATEADD(DAY, -1, GETDATE( )))) as date)
                order by [MC Surfaces, Inc].[dbo].pchord.usrnme, [MC Surfaces, Inc].[dbo].tkfprt.prtcls;
    """
    cursor.execute(sql)
    tuples = list(cursor.fetchall( ))
    data = [ ]
    for tuple in tuples:
        data.append(list(tuple))

    # Save the results to .xlsx file
    fileName = "MATERIAL ORDERED " + today + ".xlsx"
    if os.path.exists(dir + '\\Reports\\' + fileName):
        os.remove(dir + '\\Reports\\' + fileName)
    workbook = xlsxwriter.Workbook(dir + "\\Reports\\" + fileName)
    worksheet = workbook.add_worksheet( )
    worksheet.set_landscape( )
    worksheet.set_header('&C&24Weekly Material Ordered')
    worksheet.fit_to_pages(1, 0)

    # Workbook Formatting
    bold = workbook.add_format({'bold': True})
    titleRight = workbook.add_format({'bold': True, 'align': 'right'})
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    top = workbook.add_format({'valign': 'top'})
    money = workbook.add_format({'num_format': '#,##0.00'})
    total = workbook.add_format({'num_format': '#,##0.00', 'bold': True})
    quantity = workbook.add_format({'num_format': '0.00000', 'valign': 'top'})
    highlight = workbook.add_format( )
    highlight.set_pattern(1)
    highlight.set_bg_color('yellow')

    # Cell Dimensions
    worksheet.set_column("G:G", None, money)
    worksheet.set_row(0, 20)

    # Write information to the workbook
    worksheet.write('A1', 'Order #', bold)
    worksheet.write('B1', 'Vendor', bold)
    worksheet.write('C1', 'Part Class', bold)
    worksheet.write('D1', 'Part #', titleRight)
    worksheet.write('E1', 'Description', bold)
    worksheet.write('F1', 'Quantity', titleRight)
    worksheet.write('G1', '', bold)
    worksheet.write('H1', 'Total', titleRight)
    worksheet.write('I1', 'Entered', bold)
    worksheet.write('J1', 'Order Date', bold)
    worksheet.write('K1', 'User', bold)
    worksheet.write('L1', 'DEPT', bold)

    # Write Data to Worksheet
    row = 2
    col = 0
    user = ''
    userTotal = 0
    partClass = ''
    dept = ''
    sum = 0
    for on, vn, vnm, pc, cn, pn, pd, qty, ttl, ed, od, usr, prtcls in data:
        if ttl is None:
            ttl = 0
            
        if user == '':
            user = usr
            userTotal += float(ttl)
            if pc is None:
                continue
            dept = getDept(int(pc), usr)
            partClass = prtcls
        elif user != usr or dept != getDept(int(pc), usr):
            worksheet.write(row, col + 6, '$', bold)
            worksheet.write(row, col + 7, userTotal, total)
            sum += userTotal
            worksheet.write(row, col + 10, user + " Total", bold)
            user = usr
            if pc is None:
                continue
            dept = getDept(int(pc), usr)
            userTotal = float(ttl)
            row += 1
        else:
            userTotal += float(ttl)
        worksheet.write(row, col, on, top)
        worksheet.write(row, col + 1, str(vn) + " - " + vnm, top)
        worksheet.write(row, col + 2, str(pc) + " - " +  cn, top)
        worksheet.write(row, col + 3, pn, top)
        worksheet.write(row, col + 4, pd, top)
        worksheet.write(row, col + 5, qty, quantity)
        worksheet.write(row, col + 6, '$', top)
        worksheet.write(row, col + 7, float(ttl), money)
        entered = datetime.strptime(ed, '%Y-%m-%d')
        date_format = workbook.add_format({'num_format': 'mm/dd/yyyy', 'valign': 'top'})
        worksheet.write(row, col + 8, entered, date_format)
        invoice_date = datetime.strptime(od, '%Y-%m-%d')
        worksheet.write(row, col + 9, invoice_date, date_format)
        worksheet.write(row, col + 10, usr, top)
        worksheet.write(row, col + 11, getDept(int(pc), usr), top)
        row += 1
    worksheet.write(row, col + 6, '$', bold)
    worksheet.write(row, col + 7, userTotal, total)
    sum += userTotal
    worksheet.write(row, col + 10, user + " Total", bold)
    worksheet.write(row + 1, col + 6, '$', bold)
    worksheet.write(row + 1, 7, sum,  total)
    worksheet.write(row + 1, col + 10, " Grand Total", bold)

    # Conditional Formatting
    worksheet.conditional_format('C2:C' + str(row) + ')', {'type':        'cell',
                                                           'criteria':    '==',
                                                           'value':       '"2 - Inactive Billing Parts"',
                                                           'format':      highlight})

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
