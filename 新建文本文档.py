#!/usr/bin/env python
'''
@scriptName: DB2-DB2��ҵ�ӿڿ���
@function:  ����tableFile��mapping��׼�ļ����ɶ�Ӧ�Ĺ�DMP�����ļ���
@author:  ������
@createTime: 2018��11��30��

'''
import sys,os
import re
import xlrd
import win32com.client
import xlwt
import operator
import shutil

#��������ϵͳ����ӳ�䣺
#�����������ݿ�ӳ��:
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
 ��ȡԴ�ļ�
 '''
 dlist=[]

 try:
  for root,dirs,files in os.walk(filename):
   for f in files:
    print('reading',f)
    filename = ''.join((root,f))

    tb=xlrd.open_workbook(filename) #���ļ�sheet
    tSheet=tb.sheet_by_name(u'Sheet1')
    
    #Դ��Ŀ���---������Ϣ
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
     #��ȡ�ֶ���Ϣ �ֶ���Ϣ�ӵ�6��7��8��9��10�п�ʼ
     tmp=[]
     for row in range(1,nrows):
      tmp.append(str(tSheet.cell(row,col).value).strip().lstrip().rstrip())
      #print(tSheet.cell(row,col).value)
     tmpdlist.append(tmp)
    
    #��ȡ��ȡ����:
    tb_select_filt = str(tSheet.cell(1,11).value).lstrip().rstrip()
    tmpdlist.append(tb_select_filt)
    #��ȡĿ�������
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
 '''д���ݵ�tablefile
 '''
 
 #�Ա�׼ģ����и��ƣ�����һ������
 std_file= r'D:\RTC_201811\output\stdmodel\tableFile.xlsx'
 shutil.copyfile(std_file,filename)
 
 """��excel���Ĳ���"""
 xlApp=win32com.client.Dispatch('Excel.Application')
 xlApp.Visible=0
 xlApp.DisplayAlerts=0    #��̨���У�����ʾ��������
 try:
  xlBook=xlApp.Workbooks.Open(filename)
 except:
  print('���ļ�ʧ��')
  
 wt_t=xlBook.Worksheets('��')
 wt_f=xlBook.Worksheets('�ֶ�')

 tableCount=0
 tableRow=2  #�ӵڶ��п�ʼд

 for tlist in dlist:
  #[ģʽ����Ŀ��ģʽ����Ŀ�����ݿ���������������������[�ֶ�...]��[�ֶ�˵��]��[�ֶ�����...]��[pk..]��[pi...]
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
  tb_select_filt = tlist[11] #��������
  
  if len(tlist[12])<1:
   dst_table_name = tableName_en
  else:
   dst_table_name = tlist[12] #Ŀ�������
  
  if len(tb_select_filt)>1:
   mval_AP = '����'
  else:
   mval_AP = 'ȫ��'

  #��
  tmptb_row=tableCount*2
  row_source=tmptb_row+2
  row_dest=tmptb_row+3
  wt_t.Cells(row_source, colA).Value='Դ'       #Դ����ԴĿ���־
  wt_t.Cells(row_source, colB).Value=src      #Դ������������
  wt_t.Cells(row_source, colC).Value=ser_sys_mapping[src]   #Դ��ϵͳ����
  wt_t.Cells(row_source, colD).Value=ser_db2_mapping[src] #Դ�����ݿ���
  wt_t.Cells(row_source, colE).Value=ts_fromMode
  wt_t.Cells(row_source, colF).Value=tableName_en
  wt_t.Cells(row_source, colG).Value=tableName_ch
  wt_t.Cells(row_source, colH).Value='��'       #Դ��������
  wt_t.Cells(row_source, colI).Value=tableName_en                #��ʵԴ����
  wt_t.Cells(row_source, colJ).Value='��'      #Դ���Ƿ������
  wt_t.Cells(row_source, colK).Value='1'       #Դ����ҵ���
  wt_t.Cells(row_source, colL).Value='��'       #�Ƿ�ͬ��
  wt_t.Cells(row_source, colU).Value='��'       #����ɾ����־
  #���ն�ʱ:
  #wt_t.Cells(row_source, colO).Value='��'
  #wt_t.Cells(row_source, colP).Value='18'
  wt_t.Cells(row_source, colQ).Value='ÿ��'
  ######################################################
  wt_t.Cells(row_source, colAH).Value='��'     #�Ƿ�˽��
  wt_t.Cells(row_source, colAI).Value='��׼�ļ�'    #Դ��Ŀ��������� 
  wt_t.Cells(row_source, colAJ).Value='DXP_EX_DB2_TO_STD_01'  #Դ����ҵģ������ AJ
  wt_t.Cells(row_source, colAK).Value='DXP'
  wt_t.Cells(row_source, colAL).Value='LT06'
  wt_t.Cells(row_source, colAN).Value='/DW_DXP/DATA/HQ/'+ ser_sys_mapping[src]+'/#ETL_DAT#'
  wt_t.Cells(row_source, colAO).Value=ser_sys_mapping[src]+"_P_"+ src+"_" + ts_fromDataBase + "_"+ts_fromMode+"_"+tableName_en
  ######################################################

  wt_t.Cells(row_dest, colA).Value='Ŀ��'      #Ŀ�����ԴĿ���־
  wt_t.Cells(row_dest, colB).Value=dst      #Ŀ�������������
  wt_t.Cells(row_dest, colC).Value=ser_sys_mapping[dst]   #Ŀ���ϵͳ����
  wt_t.Cells(row_dest, colD).Value=ser_db2_mapping[dst]  #Ŀ������ݿ���
  wt_t.Cells(row_dest, colE).Value=ts_targetMode
  wt_t.Cells(row_dest, colF).Value= dst_table_name
  wt_t.Cells(row_dest, colG).Value=tableName_ch
  wt_t.Cells(row_dest, colH).Value='��'       # Ŀ���������
  wt_t.Cells(row_dest, colL).Value='��'       # �Ƿ�ͬ��
  wt_t.Cells(row_dest, colU).Value='��'       #����ɾ����־
  wt_t.Cells(row_dest, colAH).Value='��'     #�Ƿ�˽��
  
  #�ֶ�
  rowf=len(fieldlist)
  seqlist=[seq for seq in range(rowf)]   #�ֶ����
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
    type_len=key[8:key.find(',')]  #��ȡDECIMAL�ĳ���
    type_prc=key[key.rfind(',')+1:-1] #��ȡDECIMAL�ľ���
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
   print('δƥ���ַ�����')
  ispklist=[]
  pkseqlist=[]
  ispkCanNull=[]
  pknum=0
  for pk in pklist:
   if pk=='��':
    ispklist.append(pk)
    pkseqlist.append(pknum)
    ispkCanNull.append('��')
    pknum +=1
   else:
    ispklist.append('��')
    pkseqlist.append('')
    ispkCanNull.append('��')

  ispilist=[]
  for pi in pilist:
   if pi!='':
    ispilist.append('��')
   else:
    ispilist.append('��')

  wt_f.Cells(tableRow,colZ).Value = '�ճ���ȡ'  #��ȡ���ʹ���
  wt_f.Cells(tableRow,colAB).Value = mval_AP   #����/ȫ��
  wt_f.Cells(tableRow,colAC).Value = tb_select_filt #Ŀ���������
  wt_f.Cells(tableRow,colY).Value = tb_select_filt #��ȡ����
  wt_f.Cells(tableRow,colAD).Value = '��'    #�Ƿ�������ҵ
  
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
   
   wt_f.Cells(tmprow,colS).Value ='��'   #�����־
   wt_f.Cells(tmprow,colU).Value ='��'   #�Ƿ������ֶ�
   wt_f.Cells(tmprow,colW).Value =field
   wt_f.Cells(tmprow,colX).Value ='��׼ת��'
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
   wt_f.Cells(tmprow,colM).Value ='��'
   wt_f.Cells(tmprow,colN).Value =0
   wt_f.Cells(tmprow,colO).Value =pi
   wt_f.Cells(tmprow,colP).Value ='��'
   wt_f.Cells(tmprow,colS).Value ='��'   #�����־
   wt_f.Cells(tmprow,colU).Value ='��'   #�Ƿ������ֶ�   
   
  tableRow+=rowf*2
  tableCount+=1

 xlBook.Close(SaveChanges=1)
 del xlApp
 return

################################################################################# 
 
def writeMapFile_fun(filename,dlist):
 '''д���ݵ�mapping'''
 
 #�Ա�׼ģ����и��ƣ�����һ������
 std_file= r'D:\RTC_201811\output\stdmodel\mapping.xlsx'
 shutil.copyfile(std_file,filename)
 
 """
 �Ը��Ƶ�excel���Ĳ���
 """
 xlApp=win32com.client.Dispatch('Excel.Application')
 xlApp.Visible=0
 xlApp.DisplayAlerts=0   #��̨���У�����ʾ��������
 try:
  xlBook=xlApp.Workbooks.Open(filename)
 except:
  print('���ļ�ʧ��')
 wt_f = xlBook.Worksheets('����ӳ��')
 
 tableCount=1
 tableRow=3  #�ӵ����п�ʼдmapping
 #ÿ�������ѭ��
 for tlist in dlist:
  #��ȡԴӳ���ļ���õı���
  src = tlist[1] #Դ����������
  dst = tlist[3] #Ŀ�����������
  ts_fromMode=tlist[0]
  ts_fromDataBase=ser_db2_mapping[src]  #�޸� ser_db2_mapping
  ts_targetMode=tlist[2]
  ts_targetDataBase=ser_db2_mapping[dst]  #�޸� ser_db2_mapping
  tableName_en=tlist[4]
  fieldlist=tlist[6]
  tb_select_filt = tlist[11] #��������
  
  if len(tlist[12])<1:
   dst_table_name = tableName_en
  else:
   dst_table_name = tlist[12] #Ŀ�������
  
  #link·��
  mval_AM='/DW_DXP/DATA/HQ/'+ser_sys_mapping[src]+'/'+ser_sys_mapping[src]+'/'
  #mapping����:Դ�ļ�·��
  mval_E='/DW_DXP/DATA/HQ/'+ ser_sys_mapping[src]+'/#ETL_DAT#'
  
  if len(tb_select_filt)>1:
   mval_AP = '����'
   #mval_AQ = 'INSERT'
  else:
   mval_AP = 'ȫ��'
   
  rowf=len(fieldlist)
  tmp=tableRow
  
  #��׼�ļ���Ŀ������ӳ�䣬ÿ����ֻ��һ�У�����Ҫѭ��
  # �к��ж��Ǵ�1 ��ʼ�ġ�������Ҫ+1
  wt_f.Cells(tableRow, colAM).Value= mval_AM
  wt_f.Cells(tableRow, colAO).Value='�ճ���ȡ'
  wt_f.Cells(tableRow, colAQ).Value= 'LOAD'   #  ��������ķ�ʽ�� INSERT
  wt_f.Cells(tableRow, colAR).Value='��'
  wt_f.Cells(tableRow, colAS).Value=8
  wt_f.Cells(tableRow, colAU).Value='��'
  wt_f.Cells(tableRow, colAP).Value=mval_AP
  wt_f.Cells(tableRow, colAN).Value=tb_select_filt

  #����ӳ��  �ֶμ�ѭ��
  for (field,row) in zip(fieldlist,range(0,rowf)):
   tmprow=row+tableRow
   
   wt_f.Cells(tmprow,colA).Value ='��׼�ļ�'
   wt_f.Cells(tmprow,colB).Value ='DXP'
   wt_f.Cells(tmprow,colC).Value ='LT06'
   wt_f.Cells(tmprow,colE).Value =mval_E
   wt_f.Cells(tmprow,colF).Value =ser_sys_mapping[src]+"_P_"+ src+"_" + ts_fromDataBase + "_"+ts_fromMode+"_"+tableName_en
   wt_f.Cells(tmprow,colG).Value =field
   wt_f.Cells(tmprow,colH).Value ='��׼ת��'
   wt_f.Cells(tmprow,colI).Value ='Ŀ�����ݱ�'
   wt_f.Cells(tmprow,colJ).Value =dst       #Ŀ�����������  dst_src
   wt_f.Cells(tmprow,colK).Value =ser_sys_mapping[dst]    #Ŀ��ϵͳ����
   wt_f.Cells(tmprow,colL).Value =ser_db2_mapping[dst]     #Ŀ�������
   wt_f.Cells(tmprow,colM).Value =ts_targetMode
   wt_f.Cells(tmprow,colN).Value =dst_table_name
   wt_f.Cells(tmprow,colQ).Value =field
   wt_f.Cells(tmprow,colR).Value ='DXP_LD_STD_01_TO_DB2'
   wt_f.Cells(tmprow,colS).Value ='��'

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