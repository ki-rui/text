import xlrd
import os
import numpy as np
from config import cfg
from datetime import datetime
from logg import l
import sys
# 返回当前年份
def nyear():
    return str(datetime.now().year)

# 比较两个日期的先后
def bi(a0,a1):
    a0_l=a0.split('.')
    a1_l = a1.split('.')
    for l0,l1 in zip(a0_l,a1_l):
        if int(l0)>int(l1):
            return 1
        elif int(l0)<int(l1):
            return -1
        else: continue
    return 0

# 返回一个日期所在月的天数
def theDay0fMonth(date,add=0):
    year,month,day=date.split('.')
    year, month, day=int(year),int(month)+add,int(day)
    isRun=(year%4==0 and year%100!=0) or year%400==0
    if month==2:
        res=29 if isRun else 28
    elif month in [1,3,5,7,8,10,12]:
        res=31
    else:
        res=30
    if day>res or day <1:
        l.error('输入日期'+date+'非法！')
        sys.exit(0)
    return res

# 计算两个日期的间隔天数
def lenDate(sb,se):
    res=0
    sbtmp=sb.split('.')
    setmp = se.split('.')
    year=int(setmp[0])-int(sbtmp[0])+1
    month=int(setmp[1])-int(sbtmp[1])+1
    if year==1: # 同年日期
        if month==1: # 同月日期
            res+=int(setmp[2])-int(sbtmp[2])+1
        else:
            for i in range(int(sbtmp[1]),int(setmp[1])):
                res+=theDay0fMonth(sb,i-int(sbtmp[1]))
            res+=lenDate(sbtmp[0]+'.'+setmp[1]+'.'+sbtmp[2],se)
    else:
        res+=lenDate(sb,sbtmp[0]+'.12.31')
        res += lenDate(setmp[0] + '.1.1',se)
        for y in range(1,year-1):
            res += lenDate(str(int(sbtmp[0])+y) + '.1.1', str(int(setmp[0])+y) + '.12.31')
    return res

# 计算两个日期段的重叠天数
def isOverlap(str1,str2,con,date0):
    s1b,s1e=str1.split('-')
    s2b,s2e=str2.split('-')
    assert bi(s1b,s1e)<=0 and bi(s2b,s2e)<=0
    if bi(s1e,s2b)<0 or bi(s2e,s1b)<0:
        return 0,'',''
    elif bi(s1e,s2b)>=0 and bi(s1b,s2b)<0 and bi(s1e,s2e)<=0:
        l=lenDate(s2b,s1e)
        setContain(con,s2b,l,date0)
        return l,s2b,s1e
    elif bi(s1b,s2b)>=0 and bi(s1e,s2e)<=0:
        l = lenDate(s1b, s1e)
        setContain(con, s1b, l, date0)
        return l,s1b, s1e
    elif bi(s1b,s2b)>=0 and bi(s1b,s2e)<=0 and bi(s1e,s2e)>0:
        l = lenDate(s1b, s2e)
        setContain(con, s1b, l, date0)
        return l,s1b, s2e
    elif bi(s2b,s1b)>=0 and bi(s2e,s1e)<=0:
        l = lenDate(s2b, s2e)
        setContain(con, s2b, l, date0)
        return l,s2b, s2e
    else:
        return -1,'',''

# 根据年月日构造一个标准的日期字符串
def getDate(y,m,d):
    return str(y)+'.'+str(m)+'.'+str(d)

# 在当前日期date上增加d天，返回标准日期字符串
def addDate(date,d):
    date0 = date
    year,month,day=date0.split('.')
    year, month, day=eval(year),eval(month),eval(day)
    d0=d
    while d0>0:
        if day==theDay0fMonth(getDate(year, month, day)):
            if month==12:
                year+=1
                month=1
                day=1
            else:
                month += 1
                day = 1
            d0-=1
        if d0>=theDay0fMonth(getDate(year, month, day))-day:
            d0 -= theDay0fMonth(getDate(year, month, day)) - day
            day=theDay0fMonth(getDate(year, month, day))
        else:
            day+=d0
            d0-=d0
    return getDate(year, month, day)

# con是一个状态数组，长度等于输入时间段的天数，1代表该日已被检索到
def setContain(con,begin,l0,date0):
    index=lenDate(date0,begin)-1
    for i in range(l0):
        if con[index+i]==1:
            try:raise Exception(addDate(date0,index+i)+"被检索到了两次，这会产生严重错误！")
            except:
                l.error(addDate(date0,index+i)+"被检索到了两次，这会产生严重错误！")
                sys.exit(0)
        con[index + i] = 1
    return 1

# 如果日期字符串中包含year，则返回原始字符串；如果不包含year，则将当前年份添加至字符串头部
def addyear(date,year=nyear()):
    if len(date.split('-')[0].split('.'))==3:
        return date
    if len(date.split('-'))==2:
        ye=str(int(year)-1) if int(date.split('-')[0].split('.')[-2])>int(date.split('-')[1].split('.')[-2]) else year
        dat=ye+'.'+date.split('-')[0]+'-'+year+'.'+date.split('-')[1]
    else:
        dat=year+'.'+date
    return dat

# 查询con状态数组，全1（所有日期皆被检索到）则返回1，否则返回-1
def searchCon(con,date0):
    for i in range(con.shape[0]):
        if con[i]==0:
            try:raise Exception(addDate(date0,i)+"的记录未被找到")
            except:
                l.error(addDate(date0,i)+"的记录未被找到")
                sys.exit(0)

# 读取configs文件信息：输入时间段，文件密码
def readDate():
    datepath=os.path.join(os.path.join(cfg.root,cfg.idPath),cfg.dateFile)
    file=open(datepath,mode='r',encoding='utf-8')
    dic={}
    k=0
    for line in file:
        if line.split('\n')[0].split(';')[0].split(' ')[0]=='begin::': k=1
        elif line.split('\n')[0].split(';')[0].split(' ')[0]=='end!!!': break
        elif k==1:
            key,value=line.split(':')
            dic[key]=value.split('\n')[0].split(';')[0].split(' ')[0]
    file.close()
    return dic

# 根据输入时间段 查询命中的表格，将所有命中文件名存储在res中；返回-1代表数据文件夹不存在
def hit_xls(date,res):
    dirs=os.listdir(os.getcwd())
    if(cfg.dataPath not in dirs):
        return -1
    dates=addyear(date,nyear())

    theDay0fMonth(dates.split('-')[0])
    theDay0fMonth(dates.split('-')[1])

    l_date=lenDate(dates.split('-')[0],dates.split('-')[1])
    l.debug("程序将自动统计 "+dates+" 共 "+str(l_date)+" 天的考勤")
    con=np.zeros((l_date,))
    xls_path=os.path.join(os.getcwd(),cfg.dataPath)
    excels=os.listdir(xls_path)
    for excel in excels:
        date_tmp=addyear(excel.split('员工刷卡记录表')[0],nyear())
        l_over,over1,over2=isOverlap(dates,date_tmp,con,dates.split('-')[0])
        if l_over>0:
            res.append([excel,over1,l_over])
    searchCon(con,dates.split('-')[0])
    return 1

# 读取一个文件中的学生id信息---(学号，姓名，班级，工号，工资/工时)
def ID(id_path):
    res=[]
    table=xlrd.open_workbook(filename=id_path).sheets()[0]
    for i in range(1,table.nrows):
        ls=table.row_values(rowx=i, start_colx=0, end_colx=5)
        if ls[0]=='':continue
        if len(ls)<4:
            id=str(int(ls[0]))[2:]
            sym=0
        else:
            id=str(int(ls[3])) if ls[3]!='' else str(int(ls[0]))[2:]
            sym=1 if ls[3]!='' else 0
        if len(ls) < 5:
            salary = '0'  # 0代表工资
        else:
            salary=str(int(ls[4])) if ls[4]!='' else '0' # 0代表工资
        ls=[str(int(ls[0])),ls[1],ls[2],id,salary,sym]
        res.append(ls)
    # print(res)
    return res

# 读取id目录下所有id文件中的学生信息
def IDs():
    ids=[]
    idpath=os.path.join(cfg.root,cfg.idPath)
    dirs=os.listdir(idpath)
    for idfile in dirs:
        if idfile.split(".")[0][0:9]==cfg.idFile and idfile.split('.')[-1] in ['xls','xlsx']:
            ids.append([ID(os.path.join(idpath,idfile)),idfile.split('.')[0].split('Student')[-1]])
    return ids

conf=readDate()
hitxls=[]
hit_xls(conf['Date'],hitxls)
studID=IDs()

if __name__=='__main__':
    # main()
    s=readDate()
    print(s)