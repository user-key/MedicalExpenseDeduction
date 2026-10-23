#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#	   med.py
#	   
#	   Copyright 2020,2026 yuuzaki masahiro
#	   
#	   Released under the MIT license
#	   https://opensource.org/licenses/mit-license.php
#
    
import argparse
import configparser
import pandas as pd
from collections import namedtuple
import openpyxl
import os
import glob

def read_args():
	parser = argparse.ArgumentParser(description='医療費控除申請用一覧の作成')
	parser.add_argument("year", nargs='?', default = 'sample', help='年度(フォルダ)')
	parser.add_argument('-o', '--outfile', default = False, help="出力先 xlsxファイル")
	return parser.parse_args()
def read_initialization(name='med.ini'):
	config = configparser.ConfigParser()
	config_file = name if os.sep in name else os.path.join(os.path.dirname(os.path.abspath(__file__)),name)
	config.read(config_file)
	return config

CSVNAMES = '[0-9]*.csv'
def csv2dataframe(path, year):
	df = pd.DataFrame(columns=['name','medic','payment','refund'])
	for filename in sorted(glob.glob(os.path.join(*[path, year, CSVNAMES]))):
		s = pd.read_csv(filename, header=None, names=('date', 'pay'))
		fn = os.path.splitext(os.path.basename(filename))[0].split('_')
		ix = fn[0]
		if ix in df.index:
			print(f'duplicate index:{ix}')
			df.loc[ix] = [f'dup {ix}:{fn[1]}', fn[2], s['pay'].sum(), 0]
		else:
			df.loc[ix] = [fn[1], fn[2], s['pay'].sum(), 0]
	try:
		bf = pd.read_csv(os.path.join(*[path, year, 'refund.csv']), header=None, names=('ix', 'refund'), dtype={'ix': str, 'refund': int})
		for row in bf.itertuples():
			df.loc[row.ix]['refund'] += row.refund
	except FileNotFoundError as e:
		print('refund.csv in not found!')
	return df

def set_col(sheet, adr, value, number_format = False ,border = False):
	cell = sheet[adr]
	cell.value = value
	if number_format :
		cell.number_format = number_format 
	if border:
		cell.border = border
def set_cols(sheet, cols, number_format = False ,border = False):
	for col in cols:
	    set_col(sheet, col, cols[col], number_format, border)
def set_border(scope):
	#side = openpyxl.styles.borders.Side(style="thin", color="000000")
	side = openpyxl.styles.borders.Side(style="dotted", color="000000")
	border = openpyxl.styles.borders.Border(bottom=side)
	for rows in scope:
		for cell in rows:
			cell.border = border
def make_template(sheet, ini, year):
	col_widths = {"A":12, "B":15, "C":10, "D":12, "E":12, "F":12, "G":12}
	for col_name in col_widths:
		sheet.column_dimensions[col_name].width = col_widths[col_name]
	set_col(sheet,'A2', f'対象年:{year}')
	set_cols(sheet, {"A3":"住所:"+ini['DEFAULT']['adr'], "A4":"氏名:"+ini['DEFAULT']['name']})

CellGroup = namedtuple('CellGroup' , 'name pos')
def cell_name_merge(sheet, f, t):
	sheet[f"A{f}"].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center')
	if f != t :
		sheet.merge_cells(f"A{f}:A{t}")

def set_summary(sheet, df, pos):
	set_cols(sheet, {f'A{pos}':"名前", f'B{pos}':"医療機関", f'E{pos}':"支払金額", f'F{pos}':"保険給付等"}) 
	pos += 1
	cell_group = CellGroup(df['name'].iloc[0], pos)
	for name, medic, payment, refund in zip(df['name'],df['medic'],df['payment'],df['refund']):
		if name != cell_group.name :
			cell_name_merge(sheet, cell_group.pos, pos-1)
			cell_group = CellGroup(name, pos)
		set_cols(sheet,{f'A{pos}':name, f'B{pos}':medic})
		set_col(sheet,f'E{pos}', payment, number_format = "¥#,##0" )
		if refund:
			set_col(sheet,f'F{pos}', refund, number_format = "¥#,##0" )
		pos +=1
	cell_name_merge(sheet, cell_group.pos, pos-1)

def set_total(sheet, sum_start, size):
	set_cols(sheet,{'A6':'支払金額', 'A7':'保健給付等', 'A8':'差し引き'})
	compute_payment = f'=sum(E{sum_start+1}:E{sum_start+size})'
	compute_refund = f'=sum(F{sum_start+1}:F{sum_start+size})'
	set_cols(sheet, {'B6':compute_payment, 'B7':compute_refund, 'B8':'=B6-B7'}, number_format = "¥#,##0")
	total_pos = sum_start+size+1
	set_cols(sheet, {f'D{total_pos}':'合計', f'E{total_pos}':compute_payment, f'F{total_pos}':compute_refund}, number_format = "¥#,##0")
	set_border(sheet['B8':'B8'] ) 
	set_border(sheet[f'A{sum_start}':f'F{sum_start+size}'])
	set_border(sheet[f'D{total_pos}':f'F{total_pos}'] ) 

def main():
	args = read_args()
	ini = read_initialization()
	df = csv2dataframe(ini['DEFAULT']['path'], args.year)
	if df.empty :
		print('{} is empty!'.format(os.path.join(*[ini['DEFAULT']['path'],args.year,CSVNAMES])))
		return 1
	wb = openpyxl.Workbook()
	sh = wb.active
	make_template(sh, ini, args.year)
	sum_pos = 11
	set_summary(sh, df, sum_pos)
	set_total(sh, sum_pos, len(df))
	if args.outfile :
		wb.save(args.outfile)
	else :
		wb.save(f'{ini["DEFAULT"]["path"]}{args.year}.xlsx')
	return 0

if __name__ == '__main__':
	main()

