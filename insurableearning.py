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

	#--------- database info ----------------
	driver = r"{Microsoft Access Driver (*.mdb, *.accdb)};"

	#db_file = r"F:\\Files For\\Hai Yen Nguyen\\Practice Credits\\PC.accdb;"
	db_file = r"C:\\pycode\\CsltBC\\BC.accdb;"
	user = "admin"
	password = ""
	#--------------------------------------------------------------------
		
	print 'Process date is ' + str(time.strftime("%m/%d/%Y"))
	print 'This is the process for calculating insurable earning for cslts'
	print 'Please enter the cycle end date (mm/dd/yyyy) you want to process:'
	#-----------------------------------------------------
	#------- get cycle date ----------------------
	getcycledate = datetime.datetime.strptime(raw_input(), '%m/%d/%Y')
	endday = getcycledate

	#print 'Cycle start date is ' + str(startday)
	print 'Cycle end date is ' + str(endday)

	#-------------- get ARB Paid related to Asst and Adj -----------------	
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
	dfinsurable = dfarbpaid.loc[dfarbpaid['TransDesc'].str.contains('ASSOC|ADJ', case=False, na=False), ['Cslt','Amt']]
	#print dfinsurable.loc[dfinsurable['Cslt']==129]
	
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
	dfinsurable = dfinsurable.append(dfarbpayable, ignore_index=True)
	#print dfinsurable.loc[dfinsurable['Cslt']==129]

	#-------------- get income earning -----------------
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
			FROM tbl_Income 
			WHERE tbl_Income.SAMPLE_DATE = #''' + endday.strftime("%m/%d/%Y") + '''#;	
		'''
		
	dfr12inc = dbq.df_select(driver, db_file, sql)
	#print dfr12inc.loc[dfr12inc['Cslt']==129]	
	dfinsurable = dfinsurable.append(dfr12inc[['Cslt','Amt']], ignore_index=True)
	#print dfinsurable.loc[dfinsurable['Cslt']==129]	
	dfinsurableamt = dfinsurable.groupby(['Cslt'], as_index=False).sum()
	#print dfinsurableamt.loc[dfinsurableamt['Cslt']==129]

	#-------------- get cslt current status -----------------
	sql = '''
			SELECT DISTINCT 
				BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_SMPL_DTE AS [ReportDate]
				,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_NUM AS Cslt
				,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_NAM_FULL AS [FullName]
				,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_RO_NUM AS [RO]
				,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_SPECIAL AS [Special]
				,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_STATUS AS [CurrentStatus]
				,BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_POSITION AS [CurrentPosition]
			FROM BRANUSER_BRAN_LKG_CSLT_CURR
			INNER JOIN BRANUSER_BRAN_LKG_RO ON BRANUSER_BRAN_LKG_RO.LKG_RO_NUM = BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_RO_NUM
			WHERE BRANUSER_BRAN_LKG_CSLT_CURR.LKG_CSLT_STATUS = 'Active'
		'''

	dfresult = dbq.df_select(driver, db_file, sql)		
	
	dfresult = dfresult.merge(dfinsurableamt, on='Cslt', how='inner')
	
	print 'insurable earning process is done successfully'
		
	#--------- output to Excel ---------------------
	writer = pd.ExcelWriter('insurable.xlsx', engine='xlsxwriter')
	
	dfresult.to_excel(writer, sheet_name='earning', freeze_panes=(1,0), index=False)
	
	# Close the Pandas Excel writer and output the Excel file.
	writer.save()
	
	print 'The process is done'
	
	
	