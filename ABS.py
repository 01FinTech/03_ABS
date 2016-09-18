# -*- coding: utf-8 -*-
## ------------------------- ##
## ABS 文件
## 输入文件模版是
## Dr. Haifeng XU
## QQ: 78112407
## Email: hfxu@wxtrust.com
## modified on Sun, Sep 18, 2016
## ------------------------- ##

import numpy as np
import pandas as pd
import datetime

## 封包日
## 注意: 这里与后面的列名称保持一致
_tmp_date = '2016-10-19'
_project_fixed = pd.DataFrame ( { 'Date' : pd.Timestamp( _tmp_date) }, \
				index = {'fixDate'} )

'''
----------------------------
todo: 如何通过列名得到这些数字?
----------------------------
'''

## excel 中, 项目名称所在列
_project_name = 1

## excel 中, 项目本金所在列
_project_principle = 3

## excel 中, 项目利率所在列
_project_interest = 7

## excel 中, 派息日从该列开始
_interest_begin = 21


df = pd.read_excel( 'ABSINPUT.xlsx' )
df = df.rename( columns = { u'产品名称'         : '_col_project'        } )
df = df.rename( columns = { u'派息日期'         : '_col_interest_day'   } )
df = df.rename( columns = { u'存续金额（元）'   : '_col_principle'      } )
df = df.rename( columns = { u'融资方省'         : '_col_province'       } )
df = df.rename( columns = { u'融资方市/县'      : '_col_city'           } )
df = df.rename( columns = { u'类型'             : '_col_industry'       } )
df = df.rename( columns = { u'入池'             : '_col_include'        } )

## df_cashflow 是最终输出的现金流量表
df_cashflow = pd.DataFrame()

## --------------------------------------------------------- ##
## begin of for
## --------------------------------------------------------- ##
for i in range( len( df ) ) :

	## 创建一个空的临时 df
	df_tmp = pd.DataFrame()

	## 将该行的派息日期加入上面的临时 df
	df_tmp = pd.DataFrame( df.iloc[ i:(i+1) , _interest_begin: ] ).T

	## 为了与后面保持统一, 日期列名都是 1
	df_tmp = df_tmp.rename( columns = { i : 'Date' } )

	## 删除该列中含有空日期的行
	df_tmp = df_tmp.dropna()

	'''
        -------------------------------------
        Todo: 这里需要判断一下, 封包日是否在派息日
        -------------------------------------
	'''

	## 插入封包日
	df_tmp = df_tmp.append( _project_fixed )
	## 根据日期排序
	df_tmp = df_tmp.sort_values( by = 'Date' , ascending = [True] )

	## 插入项目名称
	df_tmp[ 'Project' ]   = df.iloc[ i ][ _project_name ]
	## 插入项目利率
	df_tmp[ 'Rate' ]      = df.iloc[ i ][ _project_interest ]
	## 插入项目本金
	df_tmp[ 'Principle' ] = df.iloc[ i ][ _project_principle ]

	## 插入相差天数
	df_tmp [ 'Day' ] = 0
	## 计算每个计息周期
	for i in range ( 1, len ( df_tmp ) ) :
		df_tmp[ 'Day' ][ i ] = ( df_tmp['Date'][i] - df_tmp['Date'][i-1]  ).days

	## --------------------------------
	## todo
	##  这里的 365 指年化天数, 应该根据具体项目修改
	## --------------------------------
	## 计算每个付息日应收利息
	df_tmp[ 'Interest' ] = df_tmp[ 'Principle' ] * df_tmp[ 'Rate' ] * df_tmp[ 'Day' ] / 365

	## 最后一期才归还本金, 其他部分还是0
	for i in range ( 0 , len ( df_tmp ) - 1 ) :
		df_tmp[ 'Principle' ][i] = 0

	## 流入现金流合计
	df_tmp[ 'Sum' ] = df_tmp[ 'Principle' ] + df_tmp[ 'Interest' ]

	'''
	## --------------------------------
	## todo
	## 如何正确显示千位符
	## 如何日期不显示分钟
	## 如何日利率用 %
	## 如何四舍五入, 保持两位小数
	## 重新命名 Index
	## 如何验证计算正确性?
	## --------------------------------
	'''

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
                .sort_values( by = 'Date' , ascending = [True] ) \
	        	.to_excel( 'ABSoutputByDate.xlsx' )

## ------------------------------------------------------------------------ ##
## 针对每一个属性进行分类统计
## ------------------------------------------------------------------------ ##
def _my_pivot_table( df , _col ) :

	df_new = pd.DataFrame()
	df_new[ u'本金' ]    	= df.groupby( _col )[ '_col_principle' ].sum()
	df_new[ u'本金占比' ]  = df_new[ u'本金' ] / df_new[ u'本金' ].sum()
	df_new[ u'个数']       = df.groupby( _col ).size()
	df_new[ u'个数占比' ]  = df_new[ u'个数' ] / df_new[ u'个数' ].sum()
	df_new = df_new.sort_values( u'本金', ascending = False )

	df_tmp = pd.DataFrame( df_new.sum(), \
				columns = { u'合计'} )

	df_new = df_new.append( df_tmp.T )
	print 'I found the df_new'
	print df_new
	return df_new
## ------------------------------------------------------------------------ ##



'''
df_province = _my_pivot_table( df, '_col_province' )
df_city     = _my_pivot_table( df, '_col_city' )
df_industry = _my_pivot_table( df, '_col_industry' )
'''


df_final = pd.DataFrame()
df_final = df_final.append( _my_pivot_table( df, '_col_province' ) )
df_final = df_final.append( _my_pivot_table( df, '_col_city' ) ) 
df_final = df_final.append( _my_pivot_table( df, '_col_industry' ) )

print df_final
df_final.to_excel( 'pivot.xlsx' )




## -------------------------------

