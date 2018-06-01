import os, re, sys, time, xlrd, pyodbc, datetime
from datetime import date
import fnmatch
import numpy as np
import pandas as pd
import itertools as it

def df_select(driver, db_file, sql):
	odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
	conn = pyodbc.connect(odbc_conn_str)
	cursor = conn.cursor()
	df = pd.read_sql_query(sql,conn)
	cursor.close()
	conn.close()
	return df
	
def get_tbldate(driver, db_file, sql):
	odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
	conn = pyodbc.connect(odbc_conn_str)
	#--------------------------------------------------------------------
	#Cdatedf = pd.read_sql_query(sql,conn)
	#latestcycledate = Cdatedf.at[(0, 'CDate')]
	cursor = conn.cursor()
	latestcycledate = cursor.execute(sql).fetchone().LDate
	cursor.close()
	conn.close()
	return latestcycledate

def update_tbldate(driver, db_file, sql):
	odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)
	conn = pyodbc.connect(odbc_conn_str)
	cursor = conn.cursor()
	cursor.execute(sql)
	cursor.commit()
	cursor.close()
	conn.close()	
	
def add_to_tbl(driver, db_file, tbl, cols, df):
	odbc_conn_str = r"DRIVER={};DBQ={};".format(driver, db_file)

	for row in df.to_records(index=False):
		values = ", ".join(['\'%s\'' % x for x in row])
		values = values.replace("'nan'", "NULL")
		
		#print values
		sql = '''INSERT INTO %s (%s) VALUES (%s);'''
		sql = sql % (tbl, cols, values)
		#print sql 
		#sys.exit("done")
		conn = pyodbc.connect(odbc_conn_str)
		cursor = conn.cursor()
		cursor.execute(sql)
		cursor.commit()
		cursor.close()
	conn.close()