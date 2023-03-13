from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
import fnmatch
import os
import PyPDF2
import re
import pathlib
import shutil

# Have to resign the link of certificate folder and the program have to look for the folder with udise name and school name  
loc_cert='D:\\PTM\\git_pt\\Pudupet\\certificate'

# Have to resing the link the excel with the student data 
loc_excel='D:\\PTM\\git_pt\\Pudupet\\Pudupet.xlsx'

# Workbook with students details
wb=load_workbook(loc_excel)
ws=wb.active

# New workbook for checked details
new=Workbook()
checked=new.active
checked['A1']='S.No.'
checked['B1']='Student Name'
checked['C1']='Father Name'
checked['D1']='Aadhar No.'
checked['E1']='Roll No.'

s_no=1
data=[]
a=ws['L2'].value
row=2

# To get the total number of active rows 
while a is not None:
	a=ws['L'+str(row)].value
	row=row+1

# Getting the students details
for i in range(2,row-1):
	s={}
	s['s_name']=ws['G'+str(i)].value.title()
	s['s_fa_name']=ws['O'+str(i)].value.title()
	s['s_aadhar']=ws['K'+str(i)].value
	s['s_roll']=ws['L'+str(i)].value
	data.append(s)

fill=PatternFill(patternType='solid',
					fgColor='00FF0000')

# Checking the details of students with the certificate
source =os.listdir(loc_cert)
for fol in source:
	loc= loc_cert+'\\'+fol
	name=fnmatch.filter(os.listdir(loc),'*_Tamil Nadu School Certificate_*.pdf')
	for f_name in name:
		file=open(loc+'\\'+f_name,'rb')
		reader=PyPDF2.PdfReader(file)
		for page in reader.pages:
			text=page.extract_text()
			pattern=re.compile(r'[0-9]{7}')
			match=pattern.findall(text)
			roll=match[0]
			for i in range(0,len(data)):
				if roll in data[i]['s_roll']:
					p_name=re.compile(data[i]['s_name'])
					p_fa_name=re.compile(data[i]['s_fa_name'])
					p_aadhar=re.compile('XXXXXXXX'+r'[0-9]{4}')
					m_name=p_name.findall(text)
					m_fa_name=p_fa_name.findall(text)
					m_aadhar=p_aadhar.findall(text)
					aadhar_c=[str(y) for y in m_aadhar[0]]
					aadhar_f=[str(x) for x in data[i]['s_aadhar']]
					print(m_name,m_fa_name,m_aadhar)
					checked['A'+str(i+2)]=s_no
					s_no += 1
					if len(m_name) > 0:
						checked['B'+str(i+2)]=m_name[0] 
					else:
						checked['B'+str(i+2)]=data[i]['s_name'] 
						checked['B'+str(i+2)].fill=fill

					if m_fa_name != None:
						checked['C'+str(i+2)]=m_fa_name[0]
					else:
						checked['C'+str(i+2)]=data[i]['s_fa_name']
						checked['B'+str(i+2)].fill=fill

					if aadhar_c[8]==aadhar_f[8] and aadhar_c[9]==aadhar_f[9] and aadhar_c[10]==aadhar_f[10] and  aadhar_c[11]==aadhar_f[11]:
						checked['D'+str(i+2)]=data[i]['s_aadhar']
					else:
						checked['D'+str(i+2)]=data[i]['s_aadhar']
						checked['B'+str(i+2)].fill=fill

					checked['E'+str(i+2)]=data[i]['s_roll']
		file.close()


new.save('Checked.xlsx')

