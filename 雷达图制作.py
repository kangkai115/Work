import numpy as np
import matplotlib.pyplot as plt
import os
import shutil    #删除文件及文件夹
import xlrd,xlwt


#收集数据
def data_collection():
    print('正在进行数据处理...')
    data = xlrd.open_workbook('1.xlsx')
    sheet = data.sheet_by_index(0)
    row = sheet.nrows
    col = sheet.ncols
    labels = sheet.row_values(0)[5:]        #去除前几列

    #每列比例计算
    max_all = []                                        #最大值列表
    col_all = []
    group_all= sheet.col_values(3)[1:]
    for c in range(5,col):
        max_data = max(sheet.col_values(c)[1:])         #每列最大值
        col_data = sheet.col_values(c)[1:]              #每列数据
        col_data = [col/max_data for col in col_data]   #每列数除以平均值
        col_data = [round(col,2) for col in col_data]   #保留两位小数
        col_all.append(col_data)
        max_all.append(max_data)
    col_all[2]=[1-col for col in col_all[2]]            #不满意反比例计算，数字越高，越不好
    col_all[2] = [round(col,2) for col in col_all[2]]   #保留两位小数

    #转化为每位员工数值
    data_all = []
    for r in range(1,row):
        data_row = sheet.row_values(r)[:5]              #每名代表基础数据
        data_row[0] = int(data_row[0])
        data_row[1] = int(data_row[1])
        for c in range(len(col_all)):
            data_row.append(col_all[c][r-1])            # 每名代表成绩数据
        data_all.append(data_row)
    print('正在进行雷达图制作,共需制作{}个雷达图'.format(row-1))
    return labels,data_all,group_all


# 创建文件夹及修改路径
def creat_folder(path):
    if os.path.exists(path):     #查找文件是否存在
        shutil.rmtree(path)      #强制删除path文件夹及里面文件
    os.mkdir(path)               #创建文件夹
    os.chdir(path)               #修改默认路径


#制作雷达图并存储
def Radar_map(labels,data_all):
    num = 1
    for data_content1 in data_all:
        data_content = data_content1[5:]
        angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False)
        data = np.concatenate((data_content, [data_content[0]]))
        angles = np.concatenate((angles, [angles[0]]))
        #画
        fig = plt.figure()
        ax = fig.add_subplot(111, polar=True) #将画布分割成1行1列，图像画在从左到右从上到下的第1块
        ax.plot(angles, data, 'ro-', linewidth=2) #ro- r红色 o为circle marker  -为实线 http://blog.csdn.net/ztf312/article/details/49933497
        ax.set_thetagrids(angles * 180 / np.pi,labels, fontproperties="SimHei")
        ax.set_title('{} {}'.format(data_content1[1],data_content1[2]), va='bottom', fontproperties="SimHei")
        # ax.grid(True)
        # plt.show()
        plt.savefig('{} {}.jpg'.format(data_content1[3],data_content1[2]))
        plt.close()                                 #关闭图片，不关闭超过20会有警报
        print('已创建{:>2}张图 {} {}.jpg'.format(num,data_content1[3],data_content1[2]))
        num+=1


#文件分类yidong+
def folder_move(group_all):
    print('正在进行文件整理....')
    group = list(set(group_all))            #根据文件名称，创建文件夹
    for g in group:
        if os.path.exists(g):
            shutil.rmtree(path)
        else :
            os.makedirs(g)

     #文件移动
    file_list = os.listdir()
    for file_name in file_list:
        if '.jpg' in file_name:                            #把“受理01组”之类的文件夹排除
            file_group = file_name.split(' ')[0]            #“受理x组 xxx”以空格分开取受理组
            shutil.move(file_name,file_group)               #移动（文件路径，目标文件路径）



def main():
    path = './雷达图'                               #定义路径
    labels,data_all,group_all=data_collection()            #数据收集
    creat_folder(path)                              #创建文件夹
    Radar_map(labels,data_all)                 #制作雷达图
    folder_move(group_all)


if __name__ == '__main__':
    main()