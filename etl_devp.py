#!/usr/bin/env python
'''
@scriptName: DB2-DB2作业接口开发
@function:  根据tableFile、mapping标准文件生成对应的供DMP导入文件。
@author:  zmb
@createTime: 2018年11月30日

'''
import sys,os
import re
import xlrd
import win32com.client
import xlwt
import operator
import shutil

#服务器和系统变慢映射：
#服务器和数据库映射:
ser_sys_mapping = {'CBH':'LN08','CBM':'LF06','CBQ':'LN08'}
ser_db2_mapping = {'CBH':'CBDMDB','CBM':'CBDMDB','CBQ':'CBMDB'}

#src_ser = 'CBQ'
#dst_src = 'CBM'

###########################################################################
colA=1
colB=2
colC=3
colD=4
colE=5
colF=6
colG=7
colH=8
colI=9
colJ=10
colK=11
colL=12
colM=13
colN=14
colO=15
colP=16
colQ=17
colR=18
colS=19
colT=20
colU=21
colV=22
colW=23
colX=24
colY=25
colZ=26
colAA=27
colAB=28
colAC=29
colAD=30
colAE=31
colAF=32
colAG=33
colAH=34
colAI=35
colAJ=36
colAK=37
colAL=38
colAM=39
colAN=40
colAO=41
colAP=42
colAQ=43
colAR=44
colAS=45
colAT=46
colAU=47

def sourceDataList_read_fun(filename):
 '''
 读取源文件
 '''
 dlist=[]

 try:
  for root,dirs,files in os.walk(filename):
   for f in files:
    print('reading',f)
    filename = ''.join((root,f))

    tb=xlrd.open_workbook(filename) #读文件sheet
    tSheet=tb.sheet_by_name(u'Sheet1')
    
    #源表，目标表---基本信息
    tbMode=str(tSheet.cell(1,0).value).strip().lstrip().rstrip()
    tbDBNm=str(tSheet.cell(1,1).value).strip().lstrip().rstrip()
    tbDestMode=str(tSheet.cell(1,2).value).strip().lstrip().rstrip()
    tbDestDBname=str(tSheet.cell(1,3).value).strip().lstrip().rstrip()
    tbEnName=str(tSheet.cell(1,4).value).strip().lstrip().rstrip()
    tbChName=str(tSheet.cell(1,5).value).strip().lstrip().rstrip()
    
    tmpdlist=[]
    tmpdlist.append(tbMode)
    tmpdlist.append(tbDBNm)
    tmpdlist.append(tbDestMode)
    tmpdlist.append(tbDestDBname)
    tmpdlist.append(tbEnName)
    tmpdlist.append(tbChName)

    nrows=tSheet.nrows

    for col in range(6,11):
     #读取字段信息 字段信息从第6，7，8，9，10行开始
     tmp=[]
     for row in range(1,nrows):
      tmp.append(str(tSheet.cell(row,col).value).strip().lstrip().rstrip())
      #print(tSheet.cell(row,col).value)
     tmpdlist.append(tmp)
    
    #读取抽取条件:
    tb_select_filt = str(tSheet.cell(1,11).value).lstrip().rstrip()
    tmpdlist.append(tb_select_filt)
    #读取目标表名称
    dst_table_name = str(tSheet.cell(1,12).value).lstrip().rstrip()
    tmpdlist.append(dst_table_name)
    
    dlist.append(tmpdlist)
 except Exception as e:
  print(e)
  dlist = []
 else:
  pass
 finally:
  pass
 return dlist
 
 
def writeTableFile_fun(filename,dlist):
 '''写数据到tablefile
 '''
 
 #对标准模板进行复制，产生一个副本
 std_file= r'D:\RTC_201811\output\stdmodel\tableFile.xlsx'
 shutil.copyfile(std_file,filename)
 
 """对excel表格的操作"""
 xlApp=win32com.client.Dispatch('Excel.Application')
 xlApp.Visible=0
 xlApp.DisplayAlerts=0    #后台运行，不显示，不警告
 try:
  xlBook=xlApp.Workbooks.Open(filename)
 except:
  print('打开文件失败')
  
 wt_t=xlBook.Worksheets('表')
 wt_f=xlBook.Worksheets('字段')

 tableCount=0
 tableRow=2  #从第二行开始写

 for tlist in dlist:
  #[模式名，目标模式名，目标数据库名，表名，表中文名，[字段...]，[字段说明]，[字段类型...]，[pk..]，[pi...]
  src=tlist[1]
  dst=tlist[3]  
  
  ts_fromMode=tlist[0]
  ts_fromDataBase=ser_db2_mapping[src]
  ts_targetMode=tlist[2]
  ts_targetDataBase=ser_db2_mapping[dst]
  tableName_en=tlist[4]
  tableName_ch=tlist[5]
  fieldlist=tlist[6]
  ftypelist=tlist[7]
  ftxtlist=tlist[8]
  pklist=tlist[9]
  pilist=tlist[10]
  tb_select_filt = tlist[11] #过滤条件
  
  if len(tlist[12])<1:
   dst_table_name = tableName_en
  else:
   dst_table_name = tlist[12] #目标表名称
  
  if len(tb_select_filt)>1:
   mval_AP = '增量'
  else:
   mval_AP = '全量'

  #表
  tmptb_row=tableCount*2
  row_source=tmptb_row+2
  row_dest=tmptb_row+3
  wt_t.Cells(row_source, colA).Value='源'       #源表：来源目标标志
  wt_t.Cells(row_source, colB).Value=src      #源表：服务器编码
  wt_t.Cells(row_source, colC).Value=ser_sys_mapping[src]   #源表：系统编码
  wt_t.Cells(row_source, colD).Value=ser_db2_mapping[src] #源表：数据库名
  wt_t.Cells(row_source, colE).Value=ts_fromMode
  wt_t.Cells(row_source, colF).Value=tableName_en
  wt_t.Cells(row_source, colG).Value=tableName_ch
  wt_t.Cells(row_source, colH).Value='表'       #源表：表类型
  wt_t.Cells(row_source, colI).Value=tableName_en                #真实源表名
  wt_t.Cells(row_source, colJ).Value='是'      #源表：是否可引用
  wt_t.Cells(row_source, colK).Value='1'       #源表：作业序号
  wt_t.Cells(row_source, colL).Value='否'       #是否同步
  wt_t.Cells(row_source, colU).Value='否'       #物理删除标志
  #日终定时:
  #wt_t.Cells(row_source, colO).Value='是'
  #wt_t.Cells(row_source, colP).Value='18'
  wt_t.Cells(row_source, colQ).Value='每日'
  ######################################################
  wt_t.Cells(row_source, colAH).Value='是'     #是否私有
  wt_t.Cells(row_source, colAI).Value='标准文件'    #源表：目标对象类型 
  wt_t.Cells(row_source, colAJ).Value='DXP_EX_DB2_TO_STD_01'  #源表：作业模板名称 AJ
  wt_t.Cells(row_source, colAK).Value='DXP'
  wt_t.Cells(row_source, colAL).Value='LT06'
  wt_t.Cells(row_source, colAN).Value='/DW_DXP/DATA/HQ/'+ ser_sys_mapping[src]+'/#ETL_DAT#'
  wt_t.Cells(row_source, colAO).Value=ser_sys_mapping[src]+"_P_"+ src+"_" + ts_fromDataBase + "_"+ts_fromMode+"_"+tableName_en
  ######################################################

  wt_t.Cells(row_dest, colA).Value='目标'      #目标表：来源目标标志
  wt_t.Cells(row_dest, colB).Value=dst      #目标表：服务器编码
  wt_t.Cells(row_dest, colC).Value=ser_sys_mapping[dst]   #目标表：系统编码
  wt_t.Cells(row_dest, colD).Value=ser_db2_mapping[dst]  #目标表：数据库名
  wt_t.Cells(row_dest, colE).Value=ts_targetMode
  wt_t.Cells(row_dest, colF).Value= dst_table_name
  wt_t.Cells(row_dest, colG).Value=tableName_ch
  wt_t.Cells(row_dest, colH).Value='表'       # 目标表：表类型
  wt_t.Cells(row_dest, colL).Value='否'       # 是否同步
  wt_t.Cells(row_dest, colU).Value='否'       #物理删除标志
  wt_t.Cells(row_dest, colAH).Value='是'     #是否私有
  
  #字段
  rowf=len(fieldlist)
  seqlist=[seq for seq in range(rowf)]   #字段序号
  typelist=[]
  flenlist=[]
  fprecision=[]
  
  for ftype in ftypelist:
   key=ftype.upper()

   if re.match(r'DATE',key):
    typelist.append('DATE')
    flenlist.append(0)
    fprecision.append(0)
    continue
    
   if re.match(r'CHARACTER\(\d+\)',key):
    typelist.append('CHARACTER')
    flenlist.append(key[10:-1])
    fprecision.append(0)
    continue
    
   if re.match(r'VARCHAR\(\d+\)',key):
    typelist.append('VARCHAR')
    flenlist.append(key[8:-1])
    fprecision.append(0)
    continue

   if re.match(r'CHAR\(\d+\)',key):
    typelist.append('CHAR')
    flenlist.append(key[5:-1])
    fprecision.append(0)
    continue

   if re.match(r'INTEGER',key):
    typelist.append('INTEGER')
    flenlist.append(0)
    fprecision.append(0)
    continue

   if re.match(r'INT',key):
    typelist.append('INTEGER')
    flenlist.append(0)
    fprecision.append(0)
    continue
   if re.match(r'BYTEINT',key):
    typelist.append('INTEGER')
    flenlist.append(0)
    fprecision.append(0)
    continue
   if re.match(r'SMALLINT',key):
    typelist.append('INTEGER')
    flenlist.append(0)
    fprecision.append(0)
    continue
   if re.match(r'TIMESTAMP',key):
    typelist.append('TIMESTAMP')
    flenlist.append(0)
    fprecision.append(6)
    continue
   if re.match(r'TIME\(\d+\)',key):
    typelist.append('TIME')
    flenlist.append(key[5:-1])
    fprecision.append(0)
    continue
   if re.match(r'DECIMAL\(\d+,\d+\)',key):
    type_len=key[8:key.find(',')]  #截取DECIMAL的长度
    type_prc=key[key.rfind(',')+1:-1] #截取DECIMAL的精度
    if int(type_len)>=24:
     type_len='24'
    if int(type_prc)>=7:
     type_prc='7'
    typelist.append('DECIMAL')
    flenlist.append(type_len)
    fprecision.append(type_prc)
    continue
   typelist.append('')
   flenlist.append('')
   fprecision.append('')
   print(tableName_en,key,end='--')
   print('未匹配字符类型')
  ispklist=[]
  pkseqlist=[]
  ispkCanNull=[]
  pknum=0
  for pk in pklist:
   if pk=='是':
    ispklist.append(pk)
    pkseqlist.append(pknum)
    ispkCanNull.append('是')
    pknum +=1
   else:
    ispklist.append('否')
    pkseqlist.append('')
    ispkCanNull.append('是')

  ispilist=[]
  for pi in pilist:
   if pi!='':
    ispilist.append('是')
   else:
    ispilist.append('否')

  wt_f.Cells(tableRow,colZ).Value = '日常抽取'  #抽取类型代码
  wt_f.Cells(tableRow,colAB).Value = mval_AP   #增量/全量
  wt_f.Cells(tableRow,colAC).Value = tb_select_filt #目标过滤条件
  wt_f.Cells(tableRow,colY).Value = tb_select_filt #抽取条件
  wt_f.Cells(tableRow,colAD).Value = '是'    #是否生成作业
  
  #print(fieldlist)#,,,,flenlist,fprecision,ispklist,pkseqlist,ispilist,ispkCanNull,range(0,rowf))

  for (field,filedtxt,filedseq,fieldtype,fieldlen,fieldpre,ispk,pkseq,pi,ispknull,row) in zip(fieldlist,ftxtlist,seqlist,typelist,flenlist,fprecision,ispklist,pkseqlist,ispilist,ispkCanNull,range(rowf)):
   tmprow=row+tableRow
   wt_f.Cells(tmprow,colA).Value =src
   wt_f.Cells(tmprow,colB).Value =ser_sys_mapping[src]
   wt_f.Cells(tmprow,colC).Value =ser_db2_mapping[src]
   wt_f.Cells(tmprow,colD).Value =ts_fromMode
   wt_f.Cells(tmprow,colE).Value =tableName_en
   wt_f.Cells(tmprow,colF).Value =field
   wt_f.Cells(tmprow,colG).Value =filedtxt
   wt_f.Cells(tmprow,colH).Value =filedseq
   wt_f.Cells(tmprow,colI).Value =fieldtype
   wt_f.Cells(tmprow,colJ).Value =fieldlen
   wt_f.Cells(tmprow,colK).Value =fieldpre
   wt_f.Cells(tmprow,colL).Value =''
   wt_f.Cells(tmprow,colM).Value =ispk
   wt_f.Cells(tmprow,colN).Value =pkseq
   wt_f.Cells(tmprow,colO).Value =pi
   wt_f.Cells(tmprow,colP).Value =ispknull
   
   wt_f.Cells(tmprow,colS).Value ='否'   #代码标志
   wt_f.Cells(tmprow,colU).Value ='否'   #是否敏感字段
   wt_f.Cells(tmprow,colW).Value =field
   wt_f.Cells(tmprow,colX).Value ='标准转换'
   #-----------------------------------------------------------------------
   tmprow=tmprow+rowf
   wt_f.Cells(tmprow,colA).Value =dst
   wt_f.Cells(tmprow,colB).Value =ser_sys_mapping[dst]
   wt_f.Cells(tmprow,colC).Value =ser_db2_mapping[dst]
   wt_f.Cells(tmprow,colD).Value =ts_targetMode
   wt_f.Cells(tmprow,colE).Value =dst_table_name
   wt_f.Cells(tmprow,colF).Value =field
   wt_f.Cells(tmprow,colG).Value =filedtxt
   wt_f.Cells(tmprow,colH).Value =filedseq
   wt_f.Cells(tmprow,colI).Value =fieldtype
   wt_f.Cells(tmprow,colJ).Value =fieldlen
   wt_f.Cells(tmprow,colK).Value =fieldpre
   wt_f.Cells(tmprow,colL).Value =''
   wt_f.Cells(tmprow,colM).Value ='否'
   wt_f.Cells(tmprow,colN).Value =0
   wt_f.Cells(tmprow,colO).Value =pi
   wt_f.Cells(tmprow,colP).Value ='是'
   wt_f.Cells(tmprow,colS).Value ='否'   #代码标志
   wt_f.Cells(tmprow,colU).Value ='否'   #是否敏感字段   
   
  tableRow+=rowf*2
  tableCount+=1

 xlBook.Close(SaveChanges=1)
 del xlApp
 return

################################################################################# 
 
def writeMapFile_fun(filename,dlist):
 '''写数据到mapping'''
 
 #对标准模板进行复制，产生一个副本
 std_file= r'D:\RTC_201811\output\stdmodel\mapping.xlsx'
 shutil.copyfile(std_file,filename)
 
 """
 对复制的excel表格的操作
 """
 xlApp=win32com.client.Dispatch('Excel.Application')
 xlApp.Visible=0
 xlApp.DisplayAlerts=0   #后台运行，不显示，不警告
 try:
  xlBook=xlApp.Workbooks.Open(filename)
 except:
  print('打开文件失败')
 wt_f = xlBook.Worksheets('数据映射')
 
 tableCount=1
 tableRow=3  #从第三行开始写mapping
 #每个表进行循环
 for tlist in dlist:
  #读取源映射文件获得的变量
  src = tlist[1] #源服务器编码
  dst = tlist[3] #目标服务器编码
  ts_fromMode=tlist[0]
  ts_fromDataBase=ser_db2_mapping[src]  #修改 ser_db2_mapping
  ts_targetMode=tlist[2]
  ts_targetDataBase=ser_db2_mapping[dst]  #修改 ser_db2_mapping
  tableName_en=tlist[4]
  fieldlist=tlist[6]
  tb_select_filt = tlist[11] #过滤条件
  
  if len(tlist[12])<1:
   dst_table_name = tableName_en
  else:
   dst_table_name = tlist[12] #目标表名称
  
  #link路径
  mval_AM='/DW_DXP/DATA/HQ/'+ser_sys_mapping[src]+'/'+ser_sys_mapping[src]+'/'
  #mapping常量:源文件路径
  mval_E='/DW_DXP/DATA/HQ/'+ ser_sys_mapping[src]+'/#ETL_DAT#'
  
  if len(tb_select_filt)>1:
   mval_AP = '增量'
   #mval_AQ = 'INSERT'
  else:
   mval_AP = '全量'
   
  rowf=len(fieldlist)
  tmp=tableRow
  
  #标准文件到目标对象的映射，每个表只有一行，不需要循环
  # 列和行都是从1 开始的。所有需要+1
  wt_f.Cells(tableRow, colAM).Value= mval_AM
  wt_f.Cells(tableRow, colAO).Value='日常抽取'
  wt_f.Cells(tableRow, colAQ).Value= 'LOAD'   #  增量插入的方式是 INSERT
  wt_f.Cells(tableRow, colAR).Value='是'
  wt_f.Cells(tableRow, colAS).Value=8
  wt_f.Cells(tableRow, colAU).Value='是'
  wt_f.Cells(tableRow, colAP).Value=mval_AP
  wt_f.Cells(tableRow, colAN).Value=tb_select_filt

  #数据映射  字段集循环
  for (field,row) in zip(fieldlist,range(0,rowf)):
   tmprow=row+tableRow
   
   wt_f.Cells(tmprow,colA).Value ='标准文件'
   wt_f.Cells(tmprow,colB).Value ='DXP'
   wt_f.Cells(tmprow,colC).Value ='LT06'
   wt_f.Cells(tmprow,colE).Value =mval_E
   wt_f.Cells(tmprow,colF).Value =ser_sys_mapping[src]+"_P_"+ src+"_" + ts_fromDataBase + "_"+ts_fromMode+"_"+tableName_en
   wt_f.Cells(tmprow,colG).Value =field
   wt_f.Cells(tmprow,colH).Value ='标准转换'
   wt_f.Cells(tmprow,colI).Value ='目标数据表'
   wt_f.Cells(tmprow,colJ).Value =dst       #目标服务器编码  dst_src
   wt_f.Cells(tmprow,colK).Value =ser_sys_mapping[dst]    #目标系统编码
   wt_f.Cells(tmprow,colL).Value =ser_db2_mapping[dst]     #目标服务器
   wt_f.Cells(tmprow,colM).Value =ts_targetMode
   wt_f.Cells(tmprow,colN).Value =dst_table_name
   wt_f.Cells(tmprow,colQ).Value =field
   wt_f.Cells(tmprow,colR).Value ='DXP_LD_STD_01_TO_DB2'
   wt_f.Cells(tmprow,colS).Value ='是'

  tableRow+=rowf
  tableCount+=1
 xlBook.Close(SaveChanges=1)
 del xlApp
 return

if __name__ == '__main__':
 import datetime
 now_dt = str(datetime.datetime.now().strftime('%Y%m%d'))
 dlist=sourceDataList_read_fun('D:/RTC_201811/input/')
 writeTableFile_fun(r'D:\RTC_201811\output\%s_tableFile_cbq_0114.xlsx'%now_dt,dlist)
 print('tablefile write done!')
 writeMapFile_fun(r'D:\RTC_201811\output\%s_mapping_cbq_0114.xlsx'%now_dt,dlist)
 print('mappingfile write done!')
