import xlrd
import os
import numpy as np

import sys
from xlutils.copy import copy
from config import cfg
import xlwt
from dataloader import lenDate,addyear,addDate,nyear,conf,hitxls,studID
from clock import clock,timeStage
from logg import l

# 根据姓名检索行号（可能返回多个值）
def searchname(table,name):
    res=[]
    for i in range(table.nrows):
        if table.cell_value(rowx=i, colx=11) == name:
            res.append(i)
    return res

# 根据工号检索行号
def searchid(table,id):
    row=table.nrows
    i=0
    j=row-1
    while i<=j:
        while table.cell_value(rowx=i, colx=1)!="工号：" and i<=j: i+=1
        while table.cell_value(rowx=j, colx=1)!="工号：" and i<=j: j-=1
        if int(table.cell_value(rowx=i, colx=3))==int(id): return i
        elif int(table.cell_value(rowx=j, colx=3))==int(id): return j
        else:
            mid=(i+j)//2
            k=-1
            m=0
            while table.cell_value(rowx=mid, colx=1)!="工号：" and mid>=i and mid<=j:
                k=1 if k==-1 else -1
                m += 1
                mid+=k*m
            if table.cell_value(rowx=mid, colx=1)!="工号：":break
            if int(table.cell_value(rowx=mid, colx=3))==int(id): return mid
            elif int(table.cell_value(rowx=mid, colx=3))>int(id):j=mid-1
            else:i=mid+1
    return -1

# 根据行号检索下面几行的考勤信息
def nextl(table,row,col_l,col_r):
    res=[]
    nex = row + 1
    while nex < table.nrows and nex < row + 100 and table.cell_value(rowx=nex, colx=1) != "工号：":
        nex += 1
    if nex < table.nrows and table.cell_value(rowx=nex, colx=1) != "工号：":
        try:
            raise Exception("从"+str(row)+"行开始查询了100行，这里出现了错误")
        except:
            l.error("从"+str(row)+"行开始查询了100行，这里出现了错误")
            sys.exit(0)
    id_row = nex - row - 2
    for i in range(id_row):
        res.append(table.row_values(rowx=row + 2 + i, start_colx=col_l, end_colx=col_r))
    return res

# 判断文件名与文件中日期的一致性
def consistent(table,path):
    ls0, ls1 = table.cell_value(rowx=2, colx=25).split('：')[-1].split('～')
    date_xl = ''
    spl = ls0.split('-')
    date_xl += spl[0] + '.' + spl[1] + '.' + spl[2]
    if not lenDate(date_xl, addyear(path.split('员工刷卡记录表')[0]).split('-')[0]) == 1:
        try: raise Exception('文件名与文件中日期不一致')
        except:
            l.error('文件名与文件中日期不一致')
            sys.exit(0)
    date_xl = ''
    spl = ls1.split('-')
    date_xl += spl[0] + '.' + spl[1] + '.' + spl[2]
    if not lenDate(date_xl, addyear(path.split('员工刷卡记录表')[0]).split('-')[1]) == 1:
        try:raise Exception('文件名与文件中日期不一致')
        except:
            l.error('文件名与文件中日期不一致')
            sys.exit(0)

# 列表元素是否全部为空
def isEmpty(ls):
    for s in ls:
        for s0 in s:
            if s0!='':return False
    return True

def extname(table,namels,col_l,col_r):
    res=[]
    if len(namels)==0:
        return -1
    for n in namels:
        table_list = nextl(table, n, col_l,col_r)
        if not isEmpty(table_list):
            res.append(table_list)
    if len(res)==1:
        return res[0]
    else:
        return -1
def mkd(root,path,filename):
    if not os.path.exists(os.path.join(root,path)):
        os.mkdir(os.path.join(root,path))
    return os.path.join(os.path.join(root,path),filename)
def salines(sheet,row,col):
    if sheet.cell_value(rowx=row,colx=col)=='':
        try:raise Exception(str(row)+'行,'+str(col)+'列 没有合计工资！')
        except:
            l.error(str(row)+'行,'+str(col)+'列 没有合计工资！')
            sys.exit(0)
    row0=row+1
    while row0<sheet.nrows and sheet.cell_value(rowx=row0,colx=col)=='':
        row0+=1
    if row0<sheet.nrows:
        return row0-row
    else:
        return sheet.nrows-row
def simple(ids):
    lenn = lenDate(addyear(conf['Date'].split('-')[0]), addyear(conf['Date'].split('-')[1]))
    da = addyear(conf['Date'], nyear())
    da = transDate(da)
    filename=ids[1] + '_' + da + '_' + cfg.outFile
    path=os.path.join(cfg.root,cfg.outPath)
    dirs=os.listdir(path)
    if not filename in dirs:
        try: raise Exception('simple计算时，输出文件不存在')
        except:
            l.error('simple计算时，输出文件不存在')
            sys.exit(0)
    wrbook = xlrd.open_workbook(os.path.join(path,filename), formatting_info=True)
    wbook = copy(wrbook)
    rsheet=wrbook.sheet_by_index(0)# 用来读
    sheet = wbook.get_sheet('考勤统计')# 用来写
    sheet_sal = wbook.get_sheet('工资')  # 用来写
    sheet_time = wbook.get_sheet('工时')  # 用来写

    s_col=cfg.startcol+3+lenn
    sal_row=2
    sal_col=cfg.startcol+4
    time_row=2
    time_col=cfg.startcol+4
    row=2
    line=salines(rsheet,row,s_col)
    k=0 # 每组学生id的索引
    style1, style2, style3 = sty()
    while(row+line<=rsheet.nrows):
        sal=0
        for i in range(lenn):
            cell=rsheet.cell_value(rowx=row+line-1,colx=cfg.startcol+3+i)
            sal+=0 if cell=='' else float(cell)
        sheet.write(row, s_col, sal,style3)
        if ids[0][k][4]=='1':
            sheet_time.write(time_row, time_col-1, sal, styh()[2])
            sheet_time.write(time_row, time_col, sal,styh()[2])
            time_row+=1
        else:
            sheet_sal.write(sal_row, sal_col-1, sal, styh()[2])
            sheet_sal.write(sal_row, sal_col, sal*cfg.hour_earn, styh()[2])
            sal_row+=1

        row+=line
        k+=1
        if row>=rsheet.nrows: break
        line=salines(rsheet,row,s_col)
    wbook.save(os.path.join(path,filename))  # 保存
    l.debug(ids[1]+'的simple计算完成！')
def styh():
    # 左栏
    style1 = xlwt.XFStyle()
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01
    # 设置背景颜色
    pattern = xlwt.Pattern()
    # 设置背景颜色的模式
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # 背景颜色
    pattern.pattern_fore_colour = 26
    style1.alignment = alignment
    # style1.pattern = pattern

    # 顶栏
    style2 = xlwt.XFStyle()
    font = xlwt.Font()
    font.height = 20 * 10
    style2.font=font
    style2.alignment = alignment
    # style2.pattern=pattern

    # 简单居中
    style3 = xlwt.XFStyle()
    style3.alignment = alignment
    return style1,style2,style3
def sty():
    # 考勤原始时间段
    style1 = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = 'Times New Roman'
    font.height = 20 * 10
    font.bold = False
    alignment = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment.vert = 0x01
    style1.font = font
    style1.alignment=alignment

    # 每人每日工时
    style2 = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = 'Times New Roman'
    font.height = 20 * 12
    font.colour_index = 61
    font.bold = True
    style2.font = font
    style2.alignment = alignment

    # 每人总工时
    style3 = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = 'Times New Roman'
    font.height = 20 * 16
    font.colour_index = 2
    font.bold = True
    style3.font = font
    style3.alignment = alignment
    return style1,style2,style3

row_sal=cfg.startrow # 全局变量，用于指示sheet_sal当前要写的行
row_time=cfg.startrow # 全局变量，用于指示sheet_time当前要写的行
def wtfile(sheet,sheet_sal,sheet_time,row,id,ls,header=False):
    global row_sal,row_time
    start_col = cfg.startcol

    style1, style2, style3=sty()
    stylel,styled,stylej=styh()
    tm=addyear(conf['Date'])
    lenn=lenDate(tm.split('-')[0],tm.split('-')[1])

    if header:
        sheet.write(row, start_col, '学号',styled)  # 学号
        sheet.write(row, start_col + 1, '姓名',styled)  # 姓名
        sheet.write(row, start_col + 2, '班级',styled)  # 班级
        for i in range(lenn):
            sheet.write(row, start_col+3+i, addDate(addyear(conf['Date'].split('-')[0]),i).replace('.','/'),styled)
            sheet.col(start_col+3+i).width = 10 * 256
        sheet.write(row, start_col + 3 + lenn, "合计工时/h",styled)
        sheet.col(start_col + 3 + lenn).width = 9 * 256

        sheet_sal.write(row_sal, start_col, '学号',stylej)  # 学号
        sheet_sal.write(row_sal, start_col + 1, '学生姓名',stylej)  # 姓名
        sheet_sal.write(row_sal, start_col + 2, '班级',stylej)  # 班级
        sheet_sal.write(row_sal, start_col + 3, '累积工作时间/h',stylej)
        sheet_sal.write(row_sal, start_col + 4, '工资',stylej)
        sheet_sal.col(start_col + 3 ).width = 15 * 256

        sheet_time.write(row_time, start_col, '学号',stylej)  # 学号
        sheet_time.write(row_time, start_col + 1, '学生姓名',stylej)  # 姓名
        sheet_time.write(row_time, start_col + 2, '班级',stylej)  # 班级
        sheet_time.write(row_time, start_col + 3, '累积工作时间/h',stylej)
        sheet_time.write(row_time, start_col + 4, '工时',stylej)
        sheet_time.col(start_col + 3).width = 15 * 256
        row_sal+=1
        row_time+=1
        return row+1
    maxx=0
    # 一个学生的所有考勤信息写入
    su=np.zeros((lenn,))
    stu_sal=0
    for table_list,col in ls:
        l=len(table_list)
        if l > maxx:
            maxx=l
        for c in range(len(table_list[0])):
            month = int(addDate(addyear(conf['Date'].split('-')[0]), col - 1 + c).split('.')[1])
            summ=0
            for n in range(len(table_list)):
                sheet.write(row+n, start_col+2+col+c, table_list[n][c].replace('\n','-'),style1)
                summ+=timeStage(table_list[n][c],isSummer=(month>=5 and month<10)).duration
            if summ!=0:
                stu_sal += summ
                su[col+c-1]=summ
    sheet.write_merge(row,row+maxx, start_col,start_col, id[0],stylel)  # 学号
    sheet.write_merge(row,row+maxx, start_col + 1,start_col + 1, id[1],stylel)  # 姓名
    sheet.write_merge(row,row+maxx, start_col + 2,start_col + 2, id[2],stylel)  # 班级
    for i in range(lenn):
        if su[i]==0: continue
        sheet.write(row+maxx, start_col + 3 + i, "{:.1f}".format(su[i]),style2)
    sheet.write_merge(row , row+maxx,start_col + 3 + lenn,start_col + 3 + lenn, stu_sal,style3)
    if id[4]=='1':
        sheet_time.write(row_time, start_col, id[0],stylej)  # 学号
        sheet_time.write(row_time, start_col + 1, id[1],stylej)  # 姓名
        sheet_time.write(row_time, start_col + 2, id[2],stylej)  # 班级
        sheet_time.write(row_time, start_col + 3, stu_sal,stylej)
        sheet_time.write(row_time, start_col + 4, stu_sal,stylej)
        row_time+=1
    else:
        sheet_sal.write(row_sal, start_col, id[0],stylej)  # 学号
        sheet_sal.write(row_sal, start_col + 1, id[1],stylej)  # 姓名
        sheet_sal.write(row_sal, start_col + 2, id[2],stylej)  # 班级
        sheet_sal.write(row_sal, start_col + 3, stu_sal,stylej)
        sheet_sal.write(row_sal, start_col + 4, stu_sal*cfg.hour_earn,stylej)
        row_sal+=1
    return row+maxx+1
def transDate(date):
    return date.replace('-','~').replace('.','-')

# 根据一组id信息读取文件
def read_xls(xls,ids):
    global row_sal,row_time
    row_sal=cfg.startrow
    row_time=cfg.startrow
    xls_path = os.path.join(cfg.root, "员工刷卡记录表")
    wrbook = xlwt.Workbook(encoding='utf-8')  # 新建工作簿
    sheet1 = wrbook.add_sheet("考勤统计")  # 新建sheet
    sheet_sal = wrbook.add_sheet("工资")  # 新建sheet
    sheet_time = wrbook.add_sheet("工时")  # 新建sheet
    wr_row = cfg.startcol
    wr_row=wtfile(sheet1,sheet_sal,sheet_time,wr_row,0,0, header=True)
    for id in ids[0]:
        ls=[]
        for xl in xls:
            path,d0,l0=xl
            path_abs=os.path.join(xls_path, path)
            index0=lenDate(addyear(path.split('-')[0]),d0)
            # filename是文件的路径名称
            workbook = xlrd.open_workbook(filename=path_abs)
            # 获取第一个sheet表格
            table = workbook.sheets()[0]
            consistent(table,path)
            ind = searchid(table, id[3])
            if ind!=-1:
                l.debug(id[1]+"的信息被发现在 "+path+" 的第"+str(ind)+"行")
                table_list=nextl(table,ind,index0,index0+l0)
                if id[5]==0 and isEmpty(table_list): # 并没有手动输入工号且根据学号查找的结果为空
                    tmp=extname(table,searchname(table, id[1]),index0,index0+l0)
                    if tmp!=-1:
                        l.warning(id[1]+"的学号检索结果为空（有可能该学生在被检索时间段内未值班），但根据其姓名在文件另一处检索到可用信息")
                        table_list=tmp
                ls.append([table_list,lenDate(addyear(conf['Date'].split('-')[0]),d0)])
            else:
                if id[5]==1: # 手动设置了工号
                    l.error(id[1] + "的工号在 " + path + " 中未被查找到，请确认您输入的工号是否准确无误！")
                    sys.exit(0)
                else:
                    tmp = extname(table, searchname(table, id[1]), index0, index0 + l0)
                    if tmp != -1:
                        l.debug(id[1] + "的工号未出现在文件中，但根据其姓名在文件另一处检索到可用信息")
                        table_list = tmp
                        ls.append([table_list,lenDate(addyear(conf['Date'].split('-')[0]),d0)])
                    else:
                        # print(id[1] + "的信息在 " + path + " 中未被找到")
                        try:raise Exception(id[1] + "的信息在 " + path + " 中未被找到")
                        except:
                            l.error(id[1] + "的信息在 " + path + " 中未被找到或者出现了重名，您最好自己动手设置工号")
                            sys.exit(0)
        wr_row = wtfile(sheet1,sheet_sal,sheet_time, wr_row, id, ls)

    da=addyear(conf['Date'],nyear())
    da=transDate(da)
    outfilename=ids[1]+'_'+da+'_'+cfg.outFile
    wrbook.save(mkd(cfg.root,cfg.outPath,outfilename))  # 保存

def rd_all(xls,allid):
    for ids in allid:
        if conf['isSimple']in ['false','f','F']: read_xls(xls, ids)
        elif conf['isSimple'] in ['true','t','T']:simple(ids)
        else:
            try:raise Exception('你输入了错误的isSimple指令')
            except:
                l.error('你输入了错误的isSimple指令')
                sys.exit(0)
    l.debug('So it\'s all done')
if __name__=='__main__':
    rd_all(hitxls,studID)
