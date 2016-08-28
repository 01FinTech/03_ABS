# -*- coding: utf-8 -*-
## ------------------------- ##
## ABS 文件
## 输入文件模版是
## Dr. Haifeng XU
## QQ: 78112407
## Email: hfxu@wxtrust.com
## modified on Mon, Aug 15, 2016
## ------------------------- ##

import numpy as np
import pandas as pd
import datetime

## 封包日
## 注意: 这里与后面的列名称保持一致
_tmp_date = '2016-08-30'
_project_fixed = pd.DataFrame ( { 'Date' : pd.Timestamp( _tmp_date) }, \
								index = {'fixDate'} )

## --------------------------------
## todo
## 如何通过列名得到这些数字?
## --------------------------------

## excel 中, 项目名称所在列
_project_name = 0

## excel 中, 项目本金所在列
_project_principle = 8

## excel 中, 项目利率所在列
_project_interest = 5

## excel 中, 派息日从该列开始
_interest_begin = 19


df = pd.read_excel ( 'ABSINPUT.xlsx' )
df = df.rename ( columns = { u'产品名称' : 'project' } )
df = df.rename ( columns = { u'派息日'  : 'interest_day' } )

print ( df )

## df_cashflow 是最终输出的现金流量表
df_cashflow = pd.DataFrame()

## --------------------------------------------------------- ##
## begin of for
## --------------------------------------------------------- ##
for i in range (0, len( df )) :

	## 创建一个空的临时 df
	df_tmp = pd.DataFrame()

	## 将该行的派息日期加入上面的临时 df
	df_tmp = pd.DataFrame( df.iloc[ i:(i+1), _interest_begin: ]).T

	## 为了与后面保持统一, 日期列名都是 1
	df_tmp = df_tmp.rename ( columns = { i : 'Date' } )

	## 删除该列中含有空日期的行
	df_tmp = df_tmp.dropna ()

	## -------------------
	## todo
	## 这里需要判断一下, 封包日是否在派息日
	## --------------------

	## --------------------

	## 插入封包日
	df_tmp = df_tmp.append( _project_fixed )
	## 根据日期排序
	df_tmp = df_tmp.sort_values( by = 'Date' )

	## 插入项目名称
	df_tmp[ 'Project' ] = df.iloc[ i ][ _project_name ]
	## 插入项目利率
	df_tmp[ 'Rate' ] = df.iloc[ i ][ _project_interest ]
	## 插入项目本金
	df_tmp[ 'Principle' ] = df.iloc[ i ][ _project_principle ]
	## 插入相差天数
	df_tmp [ 'Day' ] = 0

	## 计算每个计息周期
	for i in range ( 1, len ( df_tmp ) ) :
		df_tmp[ 'Day' ][ i ] = ( df_tmp['Date'][i] - df_tmp['Date'][i-1]  ).days
		## --------------------------------
		## 测试
		## --------------------------------
#		print ( i )
#		print (  df_tmp['Date'][i] )
#		print ( df_tmp['Date'][i-1])
#		print ( df_tmp[ 'Day' ][ i ]  )
		## --------------------------------


	## --------------------------------
	## todo
	##  这里的 365 指年化天数, 应该根据具体项目修改
	## --------------------------------
	## 计算每个付息日应收利息
	df_tmp[ 'Interest' ] =  df_tmp[ 'Principle' ] * df_tmp[ 'Rate' ] * df_tmp[ 'Day' ] / 365

	## 最后一期才归还本金, 其他部分还是0
	for i in range ( 0 , len ( df_tmp ) - 1 ) :
		df_tmp[ 'Principle' ][i] = 0

	## 流入现金流合计
	df_tmp[ 'Sum' ] = df_tmp[ 'Principle' ] + df_tmp[ 'Interest' ]

	## --------------------------------
	## todo
	## 如何正确显示千位符
	## 如何日期不显示分钟
	## 如何日利率用 %
	## 如何四舍五入, 保持两位小数
	## 重新命名 Index
	## 如何验证计算正确性?
	## --------------------------------
	##	df_tmp[ 'Sum' ] = format ( df_tmp [ 'Sum' ] , ',' )

	## 删除天数所在列
	del df_tmp[ 'Day' ]

	## 把该临时 df 追加到最后的输出中
	df_cashflow = df_cashflow.append ( df_tmp , ignore_index = True )

## --------------------------------------------------------- ##
## end of loop
## --------------------------------------------------------- ##

print ( df_cashflow )
df_cashflow.to_excel ( 'ABSoutput.xlsx' )

## 输出结果按照日期排序
## df_cashflow.sort_values( by = 'Date' ).to_excel( 'test2.xlsx' )

## 测试输出
## df_cashflow.sort( columns = ['Date'], ascending = [True] ).to_excel( 'ABSoutputByDate.xlsx')

## 按日期排序输出
## 仅仅输出封包日之后的现金流
df_cashflow[ df_cashflow['Date'] > _tmp_date ] \
	.sort( columns = ['Date'], ascending = [True] ) \
	.to_excel( 'ABSoutputByDate.xlsx')








## -------------------------------
