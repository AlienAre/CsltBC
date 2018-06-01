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

sys.path.append('C:\\pycode\\libs')
import igtools as ig
import dbquery as dbq

#------ program starting point --------	
if __name__=="__main__":
	## dd/mm/yyyy format
	print 'We will start to process annual benefit credit for both cslts and Associates.'
	print 'Process date is ' + str(time.strftime("%m/%d/%Y"))
	print 'We will work on cslts benefit credit first'
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

	db_file = r"C:\\pycode\\CsltBC\\BC.accdb;"
	user = "admin"
	password = ""
	#--------------------------------------------------------------------
		
	#--------- get list of cslts with support AL and New Business/Managed Asset ----------
	sql = '''
			SELECT 
				qry_CsltList.[RecordDate]
				,qry_CsltList.[Cslt]
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
				,BenefitCreditRate.NewBusQty
				,BenefitCreditRate.MgmtAsstRate
				,BenefitCreditRate.MgmtAsstQty
			FROM BenefitCreditRate
			WHERE (BenefitCreditRate.[Period]) = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''
	
	dfrate = dbq.df_select(driver, db_file, sql)

	dfresult = dfresult.merge(dfrate, on='SupportAL', how='left')

	dfresult.loc[dfresult['CurrentStatus'] == 'Active', 'NewBusAmt'] = dfresult['NewBus'] / dfresult['NewBusQty'] * dfresult['NewBusRate']
	dfresult.loc[dfresult['CurrentStatus'] == 'Active', 'MgmtAsstAmt'] = dfresult['MgmtAsst'] / dfresult['MgmtAsstQty'] * dfresult['MgmtAsstRate']
	dfresult.loc[dfresult['CurrentStatus'] != 'Active', ['LumpSum', 'NewBusRate', 'NewBusQty', 'MgmtAsstRate', 'MgmtAsstQty']] = np.NAN
	dfresult['BCAmt'] = dfresult['LumpSum'] + dfresult['NewBusAmt'] + dfresult['MgmtAsstAmt']

	#-------- seperate active and non-active for report use ---------------
	dfactive = dfresult.loc[dfresult['CurrentStatus'] == 'Active']
	dfnonactive = dfresult.loc[dfresult['CurrentStatus'] != 'Active']
	
	print 'cslts benefit credit process is done successfully'
	
	#-------------- Associate benefit credit part ----------------
	print 'Process date is ' + str(time.strftime("%m/%d/%Y"))
	print 'Now we work on Associates benefit credit'
	print 'Please enter the cycle end date (mm/dd/yyyy) you want to process:'
	#-----------------------------------------------------
	#------- get cycle date ----------------------
	getcycledate = datetime.datetime.strptime(raw_input(), '%m/%d/%Y')
	endday = getcycledate

	#print 'Cycle start date is ' + str(startday)
	print 'Cycle end date is ' + str(endday)
	
	#-------------- get Associate income earning -----------------
	
	#-------------- get ARB Paid related to Assoc and Adj -----------------	
	sql = '''
			SELECT 
				tARBPaid.TransDesc
				,tARBPaid.Cslt
				,tARBPaid.Amt AS [Amt]
			FROM tARBPaid
			WHERE tARBPaid.Period = #''' + endday.strftime("%m/%d/%Y") + '''#;
		'''
	dfarbpaid = dbq.df_select(driver, db_file, sql)
	#print dfarbpaid.loc[dfarbpaid['Cslt']==129]
	#dfinsurableincome = dfarbpaid.loc[dfarbpaid['TransDesc'].str.contains('ASSOC|ADJ', case=False, na=False), ['Cslt','Amt']]
	dfarbpaidassoc = dfarbpaid.loc[dfarbpaid['TransDesc'].str.contains('ASSOC|ADJ', case=False, na=False)]
	dfinsurableincome = dfarbpaidassoc[['Cslt','Amt']]
	#print dfinsurableincome.loc[dfinsurableincome['Cslt']==129]
	#-------------- get ARB Payable -----------------
	sql = '''
			SELECT 
				tARBPayable.Cslt
				,tARBPayable.Amt AS [Amt]
			FROM tARBPayable
			WHERE tARBPayable.CDate = #''' + endday.strftime("%m/%d/%Y") + '''#;
		'''
	
	dfarbpayable = dbq.df_select(driver, db_file, sql)
	#print dfarbpayable.loc[dfarbpayable['Cslt']==129]
	dfinsurableincome = dfinsurableincome.append(dfarbpayable, ignore_index=True)
	#print dfinsurableincome.loc[dfinsurableincome['Cslt']==129]

	#-------------- get Associate income earning -----------------
	sql = '''
			SELECT DISTINCT 
				tbl_Income.REP_NUM AS Cslt
				,(tbl_Income.YTD_TOTAL_CSLT_INC + tbl_Income.YTD_TOTAL_DD_INC + tbl_Income.YTD_TOTAL_RD_INC) AS YTDTotalInc
				,(tbl_Income.[YTD_MF_CMMSNS]
				+tbl_Income.[YTD_CMMSN_CRDT]
				+tbl_Income.[YTD_INS_CMMSNS]
				+tbl_Income.[YTD_SEG_FUNDS_CMMSNS]
				+tbl_Income.[YTD_OTHER_INC]
				+tbl_Income.[YTD_SEG_FUND_ARB]
				+tbl_Income.[YTD_GIF_ARB]
				+tbl_Income.[YTD_NL_ASF]
				+tbl_Income.[YTD_GIF_NL_ASF]
				+tbl_Income.[YTD_IPRO_ASF]
				+tbl_Income.[YTD_BANKING_ASF]
				+tbl_Income.[YTD_MF_NL_ASF_PREMIUM]
				+tbl_Income.[YTD_GIF_NL_ASF_PREMIUM]
				+tbl_Income.[YTD_ENH_PMT]
				+tbl_Income.[YTD_ASSOCIATE]
				+tbl_Income.[YTD_STOCKS]
				+tbl_Income.[YTD_FIXED_INCOME]
				+tbl_Income.[YTD_STRIP_BONDS]
				+tbl_Income.[YTD_IGSI_GIC]
				+tbl_Income.[YTD_FBA_ASF_EARNED]
				+tbl_Income.[YTD_SMA_ARB_EARNED]
				+tbl_Income.[YTD_SMA_NL_ASF_EARNED]
				+tbl_Income.[YTD_SALES_BONUS_PREMIUM]
				+tbl_Income.[YTD_DD_NEW_BUSINESS_INC]
				+tbl_Income.[YTD_DD_MISC_INC]
				+tbl_Income.[YTD_DD_ASSET_INC]
				+tbl_Income.[YTD_DD_BUSINESS_INC]
				+tbl_Income.[YTD_DD_RECRUITING]
				+tbl_Income.[YTD_DD_TRAINING]
				+tbl_Income.[YTD_RD_NEW_BUSINESS_INC]
				+tbl_Income.[YTD_RD_MISC_INC]
				+tbl_Income.[YTD_RD_ADMIN_SUPP]
				+tbl_Income.[YTD_RD_ASSET_INC]
				+tbl_Income.[YTD_RD_BUSINESS_INC]) AS [Amt]
				,tbl_Income.[YTD_MF_CMMSNS]
				,tbl_Income.[YTD_CMMSN_CRDT]
				,tbl_Income.[YTD_INS_CMMSNS]
				,tbl_Income.[YTD_SEG_FUNDS_CMMSNS]
				,tbl_Income.[YTD_OTHER_INC]
				,tbl_Income.[YTD_SEG_FUND_ARB]
				,tbl_Income.[YTD_GIF_ARB]
				,tbl_Income.[YTD_NL_ASF]
				,tbl_Income.[YTD_GIF_NL_ASF]
				,tbl_Income.[YTD_IPRO_ASF]
				,tbl_Income.[YTD_BANKING_ASF]
				,tbl_Income.[YTD_MF_NL_ASF_PREMIUM]
				,tbl_Income.[YTD_GIF_NL_ASF_PREMIUM]
				,tbl_Income.[YTD_ENH_PMT]
				,tbl_Income.[YTD_ASSOCIATE]
				,tbl_Income.[YTD_STOCKS]
				,tbl_Income.[YTD_FIXED_INCOME]
				,tbl_Income.[YTD_STRIP_BONDS]
				,tbl_Income.[YTD_IGSI_GIC]
				,tbl_Income.[YTD_FBA_ASF_EARNED]
				,tbl_Income.[YTD_SMA_ARB_EARNED]
				,tbl_Income.[YTD_SMA_NL_ASF_EARNED]
				,tbl_Income.[YTD_SALES_BONUS_PREMIUM]
				,tbl_Income.[YTD_DD_NEW_BUSINESS_INC]
				,tbl_Income.[YTD_DD_MISC_INC]
				,tbl_Income.[YTD_DD_ASSET_INC]
				,tbl_Income.[YTD_DD_BUSINESS_INC]
				,tbl_Income.[YTD_DD_RECRUITING]
				,tbl_Income.[YTD_DD_TRAINING]
				,tbl_Income.[YTD_RD_NEW_BUSINESS_INC]
				,tbl_Income.[YTD_RD_MISC_INC]
				,tbl_Income.[YTD_RD_ADMIN_SUPP]
				,tbl_Income.[YTD_RD_ASSET_INC]
				,tbl_Income.[YTD_RD_BUSINESS_INC]
			FROM tbl_Income 
			WHERE tbl_Income.SAMPLE_DATE = #''' + endday.strftime("%m/%d/%Y") + '''#;	
		'''
		
	dfr12inc = dbq.df_select(driver, db_file, sql)
	#print dfr12inc.loc[dfr12inc['Cslt']==129]	
	dfinsurableincome = dfinsurableincome.append(dfr12inc[['Cslt','Amt']], ignore_index=True)
	#print dfinsurableincome.loc[dfinsurableincome['Cslt']==129]	
	dfasstincome = dfinsurableincome.groupby(['Cslt'], as_index=False).sum()
	#print dfasstincome.loc[dfasstincome['Cslt']==129]

	dfasstincome = dfasstincome.merge(dfr12inc[['Cslt','YTDTotalInc']], on='Cslt', how='left') 

	#-------------- get Associate list -----------------
	sql = '''
			SELECT DISTINCT 
				BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_NUM AS Cslt
				,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_STATUS AS [Status]
				,BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_POSITION AS [Position]
			FROM BRANUSER_BRAN_LKG_CSLT
			WHERE (BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_POSITION = 'ASSOCIATE REPRESENTATIVE')
				AND (BRANUSER_BRAN_LKG_CSLT.LKG_CSLT_SMPL_DTE = #''' + endday.strftime("%m/%d/%Y") + '''#);
		'''
		
	dfasstbc = dbq.df_select(driver, db_file, sql)

	#-------------- get Associate current status -----------------
	sql = '''
			SELECT DISTINCT 
				BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_SMPL_DTE AS [RecordDate]
				,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_NUM AS Cslt
				,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_STATUS AS [CurrentStatus]
				,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_POSITION AS [CurrentPosition]
			FROM BRANUSER_BRAN_LKG_CSLT_CURR;
		'''

	dfasstcrr = dbq.df_select(driver, db_file, sql)	
	
	#-------------- get Associate benefit credit -----------------
	sql = '''
			SELECT DISTINCT 
				AsstBenefitCreditRate.[Min]
				,AsstBenefitCreditRate.[Max]
				,AsstBenefitCreditRate.[Rate]
			FROM AsstBenefitCreditRate
			WHERE AsstBenefitCreditRate.[Period] = #''' + endday.strftime("%m/%d/%Y") + '''#
		'''

	dfasstrate = dbq.df_select(driver, db_file, sql)	
	
	dfasstbc = dfasstbc.merge(dfasstcrr, on='Cslt', how='inner')
	dfasstbc = dfasstbc[['RecordDate', 'Cslt', 'Status', 'Position', 'CurrentStatus', 'CurrentPosition']]
	dfasstbc = dfasstbc.merge(dfasstincome, on='Cslt', how='left')
	
	for index, row in dfasstrate.iterrows():
		dfasstbc.loc[(dfasstbc['Amt'] >= row['Min']) & (dfasstbc['Amt'] < row['Max']), 'BC'] = row['Rate']
	
	dfasstbc.loc[dfasstbc['CurrentStatus'] != 'Active', 'BC'] = np.NAN
	
	dfactiveasstbc = dfasstbc.loc[(dfasstbc['CurrentStatus'] == 'Active') & (dfasstbc['Cslt'] < 80000)]

	dfactiveasstbcpost = dfactiveasstbc.loc[dfactiveasstbc['BC'].notnull()]
	dfactiveasstbcunpost = dfactiveasstbc.loc[dfactiveasstbc['BC'].isnull()]
	
#	writer = pd.ExcelWriter('Test.xlsx', engine='xlsxwriter')	
#	dfactiveasstbc.to_excel(writer, sheet_name='dfactiveasstbc', freeze_panes=(1,8), index=False)
#	dfactiveasstbcpost.to_excel(writer, sheet_name='dfactiveasstbcpost', freeze_panes=(1,8), index=False)
#	dfactiveasstbcunpost.to_excel(writer, sheet_name='dfactiveasstbcunpost', freeze_panes=(1,8), index=False)
#	
#	writer.save()
	
	print 'assts benefit credit process is done successfully'
		
	#--------- output to Excel ---------------------
	writer = pd.ExcelWriter('BC.xlsx', engine='xlsxwriter')
	
	dfresult.to_excel(writer, sheet_name='AllCsltBC', freeze_panes=(1,8), index=False)
	dfactive.to_excel(writer, sheet_name='ActiveCsltBC', freeze_panes=(1,8), index=False)
	dfnonactive.to_excel(writer, sheet_name='NonactiveCsltBC', freeze_panes=(1,8), index=False)
	dfasstbc.to_excel(writer, sheet_name='AllAssocBC', freeze_panes=(1,0), index=False)
	dfactiveasstbcpost.to_excel(writer, sheet_name='ActiveAssocBC', freeze_panes=(1,0), index=False)
	dfactiveasstbcunpost.to_excel(writer, sheet_name='AssocNoBC', freeze_panes=(1,0), index=False)
	dfarbpaidassoc.to_excel(writer, sheet_name='ARBPaid', freeze_panes=(1,0), index=False)
	dfarbpayable.to_excel(writer, sheet_name='ARBPayable', freeze_panes=(1,0), index=False)
	dfr12inc.to_excel(writer, sheet_name='R12Income', freeze_panes=(1,0), index=False)
	
	workbook = writer.book
	worksheet = writer.sheets['AllCsltBC']
	worksheet2 = writer.sheets['ActiveCsltBC']
	worksheet3 = writer.sheets['NonactiveCsltBC']
	worksheet4 = writer.sheets['AllAssocBC']
	worksheet5 = writer.sheets['ActiveAssocBC']
	worksheet6 = writer.sheets['AssocNoBC']
	
	# Add some cell formats.
	formatcell = workbook.add_format({'bold': True, 'align':'center'})
	formatcslt = workbook.add_format({'bold':True, 'bg_color':'#FFFF00'})
	formatdate = workbook.add_format({'num_format':'mm/dd/yyyy'})
	formatnum = workbook.add_format({'num_format':'#,##0.00'})
	formatpercent = workbook.add_format({'num_format':'0.00%'})
	formatbi = workbook.add_format({'num_format':'#,##0.00', 'bold':True, 'bg_color':'#FFFF00'})

	# Set the column width and format for AllCsltBC
	worksheet.set_column('A:A', 15, formatdate)
	worksheet.set_column('B:B', 8, formatcslt)
	worksheet.set_column('C:C', 15)
	worksheet.set_column('D:H', 14)
	worksheet.set_column('I:I', 10, formatcslt)
	worksheet.set_column('J:R', 14, formatnum)
	worksheet.set_column('S:S', 14, formatbi)

	# Set the column width and format for ActiveCsltBC
	worksheet2.set_column('A:A', 15, formatdate)
	worksheet2.set_column('B:B', 8, formatcslt)
	worksheet2.set_column('C:C', 15)
	worksheet2.set_column('D:H', 14)
	worksheet2.set_column('I:I', 10, formatcslt)
	worksheet2.set_column('J:R', 14, formatnum)
	worksheet2.set_column('S:S', 14, formatbi)
	
	# Set the column width and format for NonactiveCsltBC
	worksheet3.set_column('A:A', 15, formatdate)
	worksheet3.set_column('B:B', 8, formatcslt)
	worksheet3.set_column('C:C', 15)
	worksheet3.set_column('D:H', 14)
	worksheet3.set_column('I:I', 10, formatcslt)
	worksheet3.set_column('J:R', 14, formatnum)
	worksheet3.set_column('S:S', 14, formatbi)	
	
	# Set the column width and format for AllAssocBC
	worksheet4.set_column('A:A', 15, formatdate)
	worksheet4.set_column('B:B', 8, formatcslt)
	worksheet4.set_column('C:C', 10)
	worksheet4.set_column('D:D', 14)
	worksheet4.set_column('E:E', 10)
	worksheet4.set_column('F:F', 14)	
	worksheet4.set_column('G:G', 14, formatbi)
	worksheet4.set_column('H:H', 14, formatnum)
	worksheet4.set_column('I:I', 14, formatbi)

	# Set the column width and format for ActiveAssocBC
	worksheet5.set_column('A:A', 15, formatdate)
	worksheet5.set_column('B:B', 8, formatcslt)
	worksheet5.set_column('C:C', 10)
	worksheet5.set_column('D:D', 14)
	worksheet5.set_column('E:E', 10)
	worksheet5.set_column('F:F', 14)	
	worksheet5.set_column('G:G', 14, formatbi)
	worksheet5.set_column('H:H', 14, formatnum)
	worksheet5.set_column('I:I', 14, formatbi)

	# Set the column width and format for AssocNoBC
	worksheet6.set_column('A:A', 15, formatdate)
	worksheet6.set_column('B:B', 8, formatcslt)
	worksheet6.set_column('C:C', 10)
	worksheet6.set_column('D:D', 14)
	worksheet6.set_column('E:E', 10)
	worksheet6.set_column('F:F', 14)	
	worksheet6.set_column('G:G', 14, formatbi)
	worksheet6.set_column('H:H', 14, formatnum)
	worksheet6.set_column('I:I', 14, formatbi)
	

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()
	
	print 'The process is done'
	
	
	