#-*- coding=utf-8 -*-
import re
import quopri
import os,sys
from urllib import request
import urllib.request
import xlwt
import json
import requests

#智联
def take(x,num):
	print('正在处理第%s个....'%num)
	data1=[]	#quopri存储
	data2=[]	#gb2312存储
	dir=''
	f=open(x)
	for line in f:
		dir+=line[:-2]
	name=re.findall(r'line-height:50px">(.*?)</td></tr><tr width=3D',dir)#姓名
	data1.append(name[0])
	sex=re.findall('<font style=3D"font-weight:bold">(.*?)</font><small style=3D"colo',dir)	#性别查找，存入data1列表
	print(sex)
	if sex[0] == '=C4=D0' or sex[0] == '=C5=AE':
		data1.append(sex[0])
	else :
		data1.append(' ')
	born=re.findall('</small><font style=3D"font-weight:bold">(.*?)</font>',dir)#出生
	if  born[0][-6:]=="=D4=C2":
		data1.append(born[0])
	else:
		data1.append(born[1])
	edu=re.findall('</small>(.*?)<small ',dir)#学历
	try:
			if len(edu[1])==12:
				data1.append(edu[1])
			elif len(edu[2])==12:
				data1.append(edu[2])
			elif len(edu[3])==12:
				data1.append(edu[3])
			elif len(edu[4])==12:
				data1.append(edu[4])
	except IndexError:
			data1.append(x)
	for word in data1:    #data1列表读取
		b=quopri.decodestring(word) 	#转换 Python2此步可直接转换  pytyon3下需要下一步
		c=str(b,encoding='gb2312')
		data2.append(c)
	word_begin1 = 'ldparam3D'				#截取以此行开头字符
	word_begin2 = 'param3D'
	word_end = '">'					#截取以此行结尾字符
	f = open(x)		#打开文件
	buff = f.read()	#整页读取，因为跨行要web地址，不能readline
	buff = buff.replace('\n','')#取消换行
	buff = buff.replace('=','')#取消等号
	# print(buff)
	pat1 = re.compile(word_begin1+'(.*?)'+word_end,re.S)#截取符合条件的字符 老版匹配
	pat2 = re.compile(word_begin2+'(.*?)'+word_end,re.S)

	#对老版电话、邮箱提取
	if pat1.findall(buff):
		result = pat1.findall(buff)
		str_result=''.join(result)	#从列表中转成字符
		str_all='http://rd.zhaopin.com/resumepreview/resume/emailim?ldparam='+str_result+'&sid=121125266&site='
		f = open('d:\\10.txt', 'w')  # 创建文件以存储网页内容
		page = urllib.request.urlopen(str_all)
		data = page.read()  # 读取网页
		page1 = data.decode('utf-8')  # 转码
		try:
			f.write(page1)  # 存储网页内容
			f.close()  # 必须关闭 否则无法读取
			for line in open('d:\\10.txt'):  # 读取临时网页内容
				if re.match('                <p>电话', line):
					tel = line[59:70]
					data2.append(tel)
				if re.match('                <p>邮箱', line):
					emailadd = ''
					email = line[59:90]
					for word in email:
						if word != '<':
							emailadd = emailadd + word
						else:
							break
					data2.append(emailadd)
			return (data2)
		except UnicodeEncodeError:
			x = ['', '']
			data2.append(x)
			return (data2)
    #对新版电话、邮箱提取
	else :
		result = pat2.findall(buff)
		str_result = ''.join(result)  # 从列表中转成字符
		str_all = 'https://ihr.zhaopin.com/resumemanage/emailim.do?s=' + str_result
		f = open('d:\\办公室简历\\10.txt', 'w')  # 创建文件以存储网页内容
		headers = {'User-Agent': 'User-Agent:Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'}
		req = request.Request(str_all, headers=headers)
		page = request.urlopen(req)
		data = page.read()  # 读取网页
		page1 = data.decode('utf-8')  # 转码
		# print(page1)
		try:
			f.write(page1)  # 存储网页内容
			f.close()  # 必须关闭 否则无法读取
			for line in open('d:\\办公室简历\\10.txt'):  # 读取临时网页内容
					tel = re.compile('phone' + '(.*?)' +'email', re.S).findall(line)   #根据josn提取出来的找符合项
					tel = ''.join(tel)[3:-3]
					emailadd = re.compile('email' + '(.*?)' +'gid', re.S).findall(line)
					emailadd = ''.join(emailadd)[3:-3]
					data2.append(tel)
					data2.append(emailadd)
					# print(data2)
			return (data2)
		except UnicodeEncodeError:
			x = ['', '']
			data2.append(x)
			return(data2)


def main():
	os.chdir('d:\\办公室简历\\简历')	#修改执行路径
	file_eml=os.listdir('.')
	n=0
	for filename_eml in file_eml:	#重命名eml->txt
		n+=1
		portion=os.path.split(filename_eml)
		a=portion[1].split('.')[0]
		b=portion[1].split('.')[1]
		newfilename=a+b+'.txt'
		nfilename_txt=os.path.join('d:\\办公室简历\\简历',newfilename)
		os.rename(filename_eml,nfilename_txt)
	file_txt=os.listdir('.')
	all=[]						#****所有人员信息
	all_wrong=[]
	print('共%d个文件需要处理.....'%n)
	num=1
	for filename_txt in file_txt:	#文件夹内逐个文件执行
		try:
			s=take(filename_txt,str(num))	#调用take函数
			num+=1
			for i in s:			#每个人的信息提取存入all
				all.append(i)
		except UnicodeDecodeError:
			all_wrong.append(filename_txt)

	###excel导入###
	style=xlwt.XFStyle()
	line1=['序号','姓名','性别','出生年月','学历','电话','邮箱','备注']
	f=xlwt.Workbook()								#建立表格
	sheet1=f.add_sheet(u'sheet1',cell_overwrite_ok=True)
	alignment=xlwt.Alignment()						#创建alignment
	alignment.vert = xlwt.Alignment.VERT_CENTER 	#设置垂直对齐居中
	style.alignment=alignment						#应用alignment到sytle格式上
	line=1												#第3行
	n=0												#序号数
	sheet1.write_merge(0,0,0,7,'客服专员面试名单',style)	#合并单元格
	for l in range(0,len(line1)):
		sheet1.write(1,l,line1[l],style)			#style引用格式
	for i in range(0,len(all)):
		if i%6==0:
			sheet1.write(line+1,0,n+1,style)
			line+=1
			n+=1
		num=i%6
		sheet1.write(line,num+1,all[i],style)
	f.save('d:\\办公室简历\\客服面试名单.xls')

	with open ('d:\\办公室简历\\运行错误名单.txt','w') as f:
		for i in all_wrong:
			f.write(i)

if __name__=='__main__':
	main()
