# -*- coding:utf-8 -*-

import json
import pyodbc
import os
from os.path import join,isfile,isdir
import email
from email.header import decode_header
import datetime
import re
import email.utils
import html
from opencc import OpenCC
import jieba
import jieba.analyse
import poplib


with open ("./config.json",'r',encoding='utf-8-sig') as f:
	d= json.load(f)


#自訂義function

#######################################################################
def check_none(re):
	'''
	當正則表達式搜尋不到回傳none以避免報錯即無法辨識
	'''
	if re:
		return re.group(1)
	else:
		return None
#######################################################################
def decode_str(s):
	'''解決字符串s decode問題，
	   然而此一函數適用存在多個value,charset=decode_header(s)=[待decode內容,編碼(允許為None)]。
	   最後輸出為所有編譯字串相加
	'''
	if not s:
		return None
	value1=''
	for ele in decode_header(s):
		value,charset = ele
		if charset:
			try:
				value1+=value.decode(charset,'ignore')
			except:
				value1+=value.decode('gbk','ignore')
		else:
			value1+=str(value)
	return value1


######################################################################
def full_search(object,patterns):
	'''判斷一串文字是否有存在於一系列的pattern中
	'''
	for ele in patterns:
		rec=re.compile(ele)
		result=rec.match(object)
		if result:
			return result
			continue
######################################################################


#讀取資料庫中已掃描過之MessageID
dic_UID={}
cnxn = pyodbc.connect(d['DB'][0])
cursor = cnxn.cursor()
cursor.execute("SELECT uid FROM eml_log")
rows = cursor.fetchall()
for r in rows:
	dic_UID[r[0]]=1

#串接pop server，取得所有email的messageID
poplib._MAXLINE=20480
server='outlook.office365.com'
port=995
user='user@hotmail.com'
pswd='password'
pop=poplib.POP3_SSL(server,port)
pop.user(user)
pop.pass_(pswd)


poplist=pop.list()
numMessage = len(poplist[1])
popuid=pop.uidl()

read_uids=[]

for i in range(numMessage):
	uidl=bytes.decode(popuid[1][i])
	uid=uidl.split(' ',1)[1]
#爬取規則1:排除已爬取過並存入資料庫的Email
	if uid in dic_UID:
		continue	
	raw_email  = b"\n".join(pop.retr(i+1)[1])
	msg = email.message_from_bytes(raw_email)
#欄位1:MessageID
	MessageID= msg.get('Message-ID')

#欄位2:寄件者Eml_from
	Eml_from =str(decode_str(msg.get('from')))
	
#欄位3:收件者Eml_to
	Eml_to =str(decode_str(msg.get('to')))
	
#欄位4:寄件時間Eml_datetime
	ddt = email.utils.parsedate(msg.get('Date'))
	try:
		aaa=str(ddt[0])+'-'+str(ddt[1])+'-'+ str(ddt[2])+' '+str(ddt[3])+':'+ str(ddt[4])+':'+str(ddt[5])
		TT = datetime.datetime.strptime(aaa,"%Y-%m-%d %H:%M:%S")
		Eml_datetime=datetime.datetime.strftime(TT,"%Y-%m-%d %H:%M:%S")	 
	except:
		Eml_datetime=None
#欄位5:執行時間Run_time		
	Run_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
#欄位6:Email標題subjects
	subject = decode_str(msg.get("Subject"))
	subjects=subject
#Eml_datetime,寄件者Eml_from,收件者Eml_to,MessageID,Subject寫入eml_log資料表
	try:
		SQLlog="insert into eml_log(type,Eml_datetime,MessageID,Subject,uid,time) values(?,?,?,?,?,?)"
		cursor.execute(SQLlog,'Scanned',Eml_datetime,MessageID,subjects,uid,Run_time)
		cnxn.commit()
	except:
		print(MessageID)
#爬取規則2:指定寄件者Email
	if len(d['email_from'])!=0: 
		if not full_search(str(msg.get('from')),d['email_from']):
			continue
#欄位7:Email內容Content
	parts=[]
	for part in msg.walk():
		if part.get_filename() is not None:
			parts.append(decode_str(part.get_filename()))
		if part.get_content_type() in ['text/plain','text/html']:
				#decode_str將不適用於content的編碼解析，因其value,charset=[content,none]但能用gbk解析
			try:
				payload = part.get_payload(decode=True)
				charset = part.get_content_charset() or 'gbk'
				Rawcontent=payload.decode(charset,'ignore')
			except:
				try:
					Rawcontent=payload.decode('gbk','ignore')
				except:
					Rawcontent=None
					SQLBad="update eml_log set type ='error_file' where MessageID='%s'" %(MessageID)
					cursor.execute(SQLBad)
					cnxn.commit()
		
	if not Rawcontent:
		continue
	Content = html.unescape(re.sub('<[^>]+>','',Rawcontent))
#爬取規則3:排除沒有附件檔案的email
#	  if not parts:
#		  continue
#欄位8:詞頻(郵件關鍵字) Term Frequency
	jieba_zh='./dict.txt.big.txt'
	result = jieba.analyse.extract_tags(Content,topK=10,withWeight=True)
	TF=[]
	for tags,weight in result:
		TF.append(tags)
	TFs= ','.joing(TF)

#欄位9:另存檔名Pdf_name 

#或可call SQL function重新命名檔名
"""
	SQLFunc="select dbo.callfunctionname(?,?,?)"
	cursor.execute(SQLFunc,column1,column2,column3)
	rows=cursor.fetchall()
	cnxn.commit()
	Pdf_name_1=str(rows[0][0])
"""
#下載附件至指定資料夾
	attachments=msg.get_payload()
	for attachment in attachments:
		try:
			fname = decode_str(attachment.get_filename())
			with open(d['save'][0]+fname,'wb') as f:
				f.write(attachment.get_payload(decode=True,))
		except (Exception):
			pass
	Pdf_name=str(fname)
# 寫入email_pop資料表(每讀取一次，寫入一次，並輸出報錯的MessageID)
	try:
		SQLCommand = ("INSERT INTO email_pop(Run_time, uid, Eml_datetime, Subject, Content, Pdf_name,Eml_from,Eml_to,MessageID,TFs) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")
		cursor.execute(SQLCommand, [Run_time, uid, Eml_datetime, subjects, Content, Pdf_name,Eml_from,Eml_to,MessageID,TFs])
		cnxn.commit()
	except (pyodbc.DataError):
		print(MessageID)
		
#獲取寫入資料庫的email其uid
	read_uids.append(uid)

pop.quit()


#取得當下pop3中email排序index、對應uid
pop=poplib.POP3(server,port)
pop.user(user)
pop.pass_(pswd)

popuids=pop.uidl()
uids={}
for u in popuid[1]:
	ustr=bytes.decode(u)
	uids[ustr.split(' ',1)[1]]=ustr.split(' ',1)[0]
#獲取應刪除email之對應index
index_del=[]
for ele in read_uids:
	try:
		index_del.append(uids[ele])
	except:
		pass
#最後整體刪除所有index對應之email
for index in index_del:
	pop.dele(index)
pop.quit()

