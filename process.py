#--------------------------------------------
#version:		1.0.0.1
#author:		West
#Description:	used to prepare annual consultant benefit credit
#Assumptions:	report end date is 03/31 each year
#				New Business is past 12 trailing months sales credits/new business 
# 
#--------------------------------------------


import os, re, sys, time, xlrd, pyodbc, datetime
from datetime import date
import fnmatch
import numpy as np
import pandas as pd
import itertools as it
from openpyxl import load_workbook
from shutil import copyfile
import myfun as dd
import dbquery as dbq

#------ program starting point --------	
if __name__=="__main__":
	## dd/mm/yyyy format
	print 'Process date is ' + str(time.strftime("%m/%d/%Y"))
	print 'Please enter the cycle end date (mm/dd/yyyy) you want to process:'
	#-----------------------------------------------------
	#------- get cycle date ----------------------
	getcycledate = datetime.datetime.strptime(raw_input(), '%m/%d/%Y')
	endday = getcycledate
	#startday = datetime.datetime.strptime('1/1/' + str(endday.year), '%m/%d/%Y')
	supportalyear = endday.year

	#print 'Cycle start date is ' + str(startday)
	print 'Cycle end date is ' + str(endday)
	#--------- database info ----------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"

	#db_file = r"F:\\Files For\\Hai Yen Nguyen\\Practice Credits\\PC.accdb;"
	db_file = r"C:\\pycode\\CsltBC\\BC.accdb;"
	user = "admin"
	password = ""
	#--------------------------------------------------------------------

	#-------------- get assistant info and amount -----------------
	sql = '''
			SELECT DISTINCT
				qry_CsltwR12Inc.Cslt
				,qry_CsltwR12Inc.[Status]
				,qry_CsltwR12Inc.[Position]
				,qry_CsltwR12Inc.CurrentStatus
				,qry_CsltwR12Inc.CurrentPosition
				,qry_CsltwR12Inc.YTDTotalInc	
			FROM qry_CsltwR12Inc
			WHERE (qry_CsltwR12Inc.[Position] LIKE "*ASSOCIATE REPRESENTATIVE*")
				AND (qry_CsltwR12Inc.[CDate] = #''' + endday.strftime("%m/%d/%Y") + '''#);
		'''

	dfasstbc = dbq.df_select(driver, db_file, sql)
	
	#-------------- get assistant benefit credit -----------------
	sql = '''
			SELECT DISTINCT 
				AsstBenefitCreditRate.[Min]
				,AsstBenefitCreditRate.[Max]
				,AsstBenefitCreditRate.[Rate]
			FROM AsstBenefitCreditRate
			WHERE AsstBenefitCreditRate.[Period] = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfasstrate = dbq.df_select(driver, db_file, sql)	
	print dfasstrate
	
	for index, row in dfasstrate.iterrow:
		dfasstbc.loc[dfasstbc['YTDTotalInc'] >= row['Min'] & dfasstbc['YTDTotalInc'] < row['Max'], 'BC'] = row['Rate']
	print dfasstbc.head()
	sys.exit('--------stop---------')
		

	
	#--------- get list of cslts with support AL and New Business/Managed Asset ----------
	sql = '''
			SELECT 
				qry_CsltList.[Cslt]
				,qry_CsltList.[Name]
				,qry_CsltList.[Status]
				,qry_CsltList.[Position]
				,qry_CsltList.[CurrentStatus]
				,qry_CsltList.[CurrentPosition]
				,qry_CsltList.[TermDate]
				,qry_CsltList.[SupportAL]
				,qry_CsltList.[NewBus]
				,qry_CsltList.[MgmtAsst]
			FROM qry_CsltList
			WHERE (qry_CsltList.[EYear]) = ''' + str(supportalyear) + ''' AND qry_CsltList.LKG_CSLT_SMPL_DTE = #''' + endday.strftime("%m/%d/%Y") + '''#
			ORDER BY 	
				qry_CsltList.[Cslt]
		'''

	dfresult = dbq.df_select(driver, db_file, sql)
	dfresult['SupportAL'] = dfresult['SupportAL'].values.astype(np.int64)

	#----------- get current year rate for benefit credit ------
	sql = '''
			SELECT BenefitCreditRate.AL AS [SupportAL]
				,BenefitCreditRate.LumpSum
				,BenefitCreditRate.NewBusRate
				,BenefitCreditRate.NewBusQty9
				,BenefitCreditRate.MgmtAsstRate
				,BenefitCreditRate.MgmtAsstQty
			FROM BenefitCreditRate
			WHERE ([BenefitCreditRate].[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''
	
	dfrate = dbq.df_select(driver, db_file, sql)

	dfresult = dfresult.merge(dfrate, on='SupportAL', how='left')

	dfresult.loc[dfresult['CurrentStatus'] == 'Active', 'NewBusAmt'] = dfresult['NewBus'] / dfresult['NewBusQty'] * dfresult['NewBusRate']
	dfresult.loc[dfresult['CurrentStatus'] == 'Active', 'MgmtAsstAmt'] = dfresult['MgmtAsst'] / dfresult['MgmtAsstQty'] * dfresult['MgmtAsstRate']
	dfresult.loc[dfresult['CurrentStatus'] != 'Active', ['LumpSum', 'NewBusRate', 'NewBusQty', 'MgmtAsstRate', 'MgmtAsstQty']] = np.NAN
	dfresult['BCAmt'] = dfresult['LumpSum'] + dfresult['NewBusAmt'] + dfresult['MgmtAsstAmt']


		
	
	#--------- output to Excel ---------------------
	writer = pd.ExcelWriter('BC.xlsx', engine='xlsxwriter')
	
	dfresult.to_excel(writer, sheet_name='BC', startrow=1, freeze_panes=(2,8), index=False)

	workbook = writer.book
	worksheet = writer.sheets['BC']
	
	# Add some cell formats.
	formatcell = workbook.add_format({'bold': True, 'align':'center'})
	formatcslt = workbook.add_format({'bold':True, 'bg_color':'#FFFF00'})
	formatdate = workbook.add_format({'num_format':'mm/dd/yyyy'})
	formatnum = workbook.add_format({'num_format':'#,##0.00'})
	formatpercent = workbook.add_format({'num_format':'0.00%'})
	formatbi = workbook.add_format({'num_format':'#,##0.00', 'bold':True, 'bg_color':'#FFFF00'})

	# Set the column width and format
	worksheet.set_column('A:A', 8, formatcslt)
	worksheet.set_column('B:B', 15)
	worksheet.set_column('C:F', 14)
	worksheet.set_column('H:H', 10, formatcslt)
	worksheet.set_column('I:Q', 14, formatnum)
	worksheet.set_column('R:R', 14, formatbi)

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()
	
	print 'The process is done'
	
	
	