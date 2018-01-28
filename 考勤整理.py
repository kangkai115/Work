#-*- coding:utf-8 -*-
import xlrd,xlwt
from xlrd import xldate_as_datetime
import os
import shutil    #删除文件及文件夹



def read():
    data = xlrd.open_workbook('考勤.xlsx')
    sheet1 = data.sheet_by_index(0) #部门反馈的应到数据放在sheet1
    sheet2 = data.sheet_by_index(1) #部门提供的未刷脸原因
    sheet3 = data.sheet_by_index(2) #考勤机数据放sheet3
    rows1 = sheet1.nrows
    rows2 = sheet2.nrows
    rows3 = sheet3.nrows
    sheet1_all = []
    sheet2_all = []
    sheet3_all = []

    for i in range(1,rows1):
        s1 = sheet1.row_values(i)
        s1[3] = xldate_as_datetime(s1[3],0) #将a[0]日期从数字格式编程datetime格式
        s1[3] = s1[3].strftime('%Y-%m-%d')  #将datetime格式的a[0]编程日期
        s1[2] = int(s1[2]) #工号编程整数
        s1[4] = xldate_as_datetime(s1[4],0).strftime('%H:%M:%S') #修改上班时间格式
        s1[5] = xldate_as_datetime(s1[5], 0).strftime('%H:%M:%S')#修改下班时间格式
        sheet1_all.append(s1)

    for j in range(1,rows2):
        s2 = sheet2.row_values(j)
        s2[1] = int(s2[1]) #工号编程整数
        s2[2] = xldate_as_datetime(s2[2],0).strftime('%Y-%m-%d') #修改上班时间格式
        sheet2_all.append(s2)

    for k in range(1,rows3):
        s3 = sheet3.row_values(k)
        del s3[2]         #删除列表中第二个为元素，该元素excel中为空
        if s3[4] != ''and s3[5] != '': #当上下班都有刷脸记录
            s3[0] = int(s3[0])
            s3[3] = xldate_as_datetime(s3[3], 0).strftime('%Y-%m-%d')
            s3[4] = xldate_as_datetime(s3[4],0).strftime('%H:%M:%S') #修改上班刷脸时间格式
            s3[5] = xldate_as_datetime(s3[5], 0).strftime('%H:%M:%S')#修改下班刷脸时间格式
        elif s3[4] == '':           #仅有下班
            s3[0] = int(s3[0])
            s3[3] = xldate_as_datetime(s3[3], 0).strftime('%Y-%m-%d')
            s3[5] = xldate_as_datetime(s3[5], 0).strftime('%H:%M:%S')  # 修改下班刷脸时间格式
        else:
            s3[0] = int(s3[0])      #仅有上班
            s3[3] = xldate_as_datetime(s3[3], 0).strftime('%Y-%m-%d')
            s3[4] = xldate_as_datetime(s3[4], 0).strftime('%H:%M:%S')  # 修改上班刷脸时间格式
        sheet3_all.append(s3)

    return sheet1_all,sheet2_all,sheet3_all,rows1,rows2,rows3



#应到和考勤比对
def compare(sheet1_all,sheet3_all):
    compare_all = [] #对比结果列表，最后导入excel
    for i in range(0,len(sheet1_all)):
        for j in range (0,len(sheet3_all)-1):
            compare = []
            #如果两表格日期，时间都一样，汇总
            if sheet1_all[i][3] == sheet3_all[j][3]  and sheet1_all[i][2] == sheet3_all[j][0]:
                compare.extend(sheet1_all[i][0:5])
                compare.append(sheet3_all[j][4])
                compare.append('')
                compare.append(sheet1_all[i][5])
                compare.append(sheet3_all[j][5])
                compare.extend(['','','',''])
                # 上午刷脸核对
                if compare[5] == '':
                    compare[6] = '未刷脸'
                elif compare[4] >= compare[5]:
                    compare[6] = '正常'
                else:
                    compare[6] = '迟到'
                # 下午刷脸核对
                if compare[8] == '':
                    compare[9] = '未刷脸'
                elif compare[7] <= compare[8]:
                    compare[9] = '正常'
                else:
                    compare[9] = '早退'
                #总体核对
                if compare[6] == '正常' and compare[9] == '正常':
                    compare[10] = '正常'
                else:
                    compare[10] = '异常'
                compare_all.append(compare)
                break
            # 整天没刷脸的通过下面找出，当应到数据在实到中没数据，且实到都找完还没匹配上，则将应到加入对比总表，时间为空
            elif j == len(sheet3_all)-2:
                compare.extend(sheet1_all[i][0:5])
                compare.extend(['', '', ])
                compare.append(sheet1_all[i][5])
                compare.extend(['', '', '', '',''])
                # 上午刷脸核对
                if compare[5] == '':
                    compare[6] = '未刷脸'
                elif compare[4] >= compare[5]:
                    compare[6] = '正常'
                else:
                    compare[6] = '迟到'
                # 下午刷脸核对
                if compare[8] == '':
                    compare[9] = '未刷脸'
                elif compare[7] <= compare[8]:
                    compare[6] = '正常'
                else:
                    compare[6] = '早退'
                #总体核对
                if compare[6] == '正常' and compare[9] == '正常':
                    compare[10] = '正常'
                else:
                    compare[10] = '异常'
                compare_all.append(compare)

    return compare_all


#部门提交的未刷脸原因和考勤结果比对
def abnormal_compare(compare_all,sheet2_all):
    for i in range(0,len(compare_all)):
        for j in range(0,len(sheet2_all)):
            if compare_all[i][2] == sheet2_all[j][1] and compare_all[i][3] == sheet2_all[j][2]:
                if sheet2_all[j][3] == '全天':
                    compare_all[i][11] = '全天'
                    compare_all[i][12] = '正常'
                elif sheet2_all[j][3] == '上午':
                    if compare_all[i][8] == '正常':
                        compare_all[i][11] = '上午'
                        compare_all[i][12] = '正常'
                    else :
                        compare_all[i][11] = '上午'
                        compare_all[i][12] = '异常'
                else:
                    if compare_all[i][6] == '正常':
                        compare_all[i][11] = '上午'
                        compare_all[i][12] = '正常'
                    else :
                        compare_all[i][11] = '上午'
                        compare_all[i][12] = '异常'

    return compare_all


#找出最终结果是异常的给abnormal_all
def abnormal(compare_all):
    abnormal_all = []
    for i in range(0,len(compare_all)):
        if compare_all[i][10] == '异常' and compare_all[i][12] != '正常' :
            abnormal_all.append(compare_all[i])
    return abnormal_all


#分部门保存在一个list里 一个部门嵌套一个list
def department(abnormal_all):
    department_name = []
    department_all = []

    for i in range(0,len(abnormal_all)):
        department_name.append(abnormal_all[i][0])
    department_name=list(set(department_name))

    for i in range(0,len(department_name)):
        department = []
        for j in range(0,len(abnormal_all)):
            if abnormal_all[j][0] == department_name[i]:
                department.append(abnormal_all[j])
        department_all.append(department)

    return department_all



def write_all(compare_all_new,abnormal_all_new):
    row1 = len(compare_all_new)
    col1 = len(compare_all_new[0])
    row2 = len(abnormal_all_new)
    col2 = len(abnormal_all_new[0])
    title = ['部门', '姓名', '工号', '日期', '上班时间', '上班刷脸','上班比对', '下班时间', '下班刷脸','下班比对','总体情况','部门上报','最终结果']
    f = xlwt.Workbook()  # 建立表格
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)

    for i in range(0, len(title)):
        sheet1.write(0,i,title[i])
    for r in range(1,row1+1):
        for c in range(0,col1):
            sheet1.write(r,c,compare_all_new[r-1][c])
    sheet2 = f.add_sheet(u'sheet2', cell_overwrite_ok=True)

    for i in range(0,len(title)):
        sheet2.write(0,i,title[i])
    for r in range(1,row2+1):
        for c in range(0, col2):
            sheet2.write(r,c,abnormal_all_new[r-1][c])
    f.save('.考勤结果.xls')



def creat_folder(path):
    if os.path.exists(path):      #查找文件是否存在
        shutil.rmtree(path)       #强制删除path文件夹及里面文件
    os.mkdir(path)


def write_part(department_all):
    title = ['部门', '姓名', '工号', '日期', '上班时间', '上班刷脸', '上班比对', '下班时间', '下班刷脸', '下班比对', '总体情况', '部门上报', '最终结果']
    row = len(department_all)

    for i in range(0,row):
        f = xlwt.Workbook()  # 建立表格
        sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
        for j in range(0, len(title)):
            sheet1.write(0, j, title[j])
        for r in range(1, len(department_all[i])+1):
            for c in range(0, len(department_all[i][r-1])):
                sheet1.write(r, c, department_all[i][r-1][c])
        f.save('./各部门/{}.xls'.format(department_all[i][r-1][0]))



def main():
    sheet1_all,sheet2_all,sheet3_all,rows1,row2,rows3=read()   #提取原始表格中的三页值
    compare_all = compare(sheet1_all,sheet3_all)               #对1,3页值对比（应到和考勤机）
    compare_all_new = abnormal_compare(compare_all,sheet2_all) #将1,3页比对结果 和 2页比较（对比出没刷脸部门也没提交理由的）
    abnormal_all_new = abnormal(compare_all_new)               #将状态异常的都提取出来
    department_all=department(abnormal_all_new)                #将状态异常的按部门提取出来
    write_all(compare_all_new,abnormal_all_new)                #总表写入
    creat_folder('./各部门')
    write_part(department_all)                                 #各部门表写入



if __name__=='__main__':
    main()