import pyodbc
import datetime
import os
import config
import WeeklyWarrantyBC
import WeeklyWarrantyChargeable
import WeeklyWarrantyLoss
import WeeklyMaterialOrdered
import WeeklyCompletions
import WeeklyTemplateLoss
import WeeklyRevenueReport
import WeeklyWarrantyTechRepairs
import WeeklyGraniteReports

# Current Date for file name
today = datetime.date.today( ).strftime("%m_%d_%y")

# Get Current Directory
dir = os.getcwd( )

WeeklyWarrantyBC.runQuery(today, dir, config.dsn)
WeeklyWarrantyChargeable.runQuery(today, dir, config.dsn)
WeeklyWarrantyLoss.runQuery(today, dir, config.dsn)
# WeeklyMaterialOrdered.runQuery(today, dir, config.dsn)
WeeklyCompletions.runQuery(today, dir, config.dsn)
WeeklyTemplateLoss.runQuery(today, dir, config.dsn)
WeeklyRevenueReport.runQuery(dir, config.dsn)
WeeklyWarrantyTechRepairs.runQuery(today, dir, config.dsn)
WeeklyGraniteReports.runQuery(today, dir, config.dsn)
