# -*- coding: utf-8 -*- 
import xlrd, xlwt
import datetime
import calendar
import time

CUSTOM_HOLIDAY = []         # 自定义本月假期
TOTAL_TIME = 390            # 每天上班时长（单位：min）

MOR_BEGIN1 = "8:15:59"      # 上午上班时间
MOR_BEGIN2 = "9:00:59"

MOR_END1 = "11:30:59"       # 上午下班时间
MOR_END2 = "12:20:59" 

AFT_BEGIN1 = "13:00:59"     # 下午上班时间
AFT_BEGIN2 = "14:00:59"     

AFT_END1 = "17:30:59"       # 下午下班时间
AFT_END2 = "18:20:59" 

ADD_NAME = u"李毅"          # 增加记录的人名
                            # 增加记录的时间
ADD_TIME = ["2014-12-11 8:30:01", "2014-12-12 8:30:02"]  

def open_excel(filename):
    try:
        data = xlrd.open_workbook(filename)
        return data
    except Exception, e:
        print str(e)

# 根据索引获取Excel表格中的数据  参数:filename：Excel文件路径, colnameIndex：表头列名所在行的索引，byIndex：表的索引
def excel_table_byindex(filename = 'file.xls', colnameIndex = 0, byIndex = 0):
    data = open_excel(filename)
    table = data.sheets()[byIndex]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    dic = {}
    lst = []
    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            name = row[1]
            if not dic.has_key(name):
                dic[name] = []
            time = datetime.datetime.strptime(row[3], "%Y-%m-%d %H:%M:%S")
            dic[name].append(time)
            lst.append(row)

    return lst, dic, table

def workday(year, month, custom = CUSTOM_HOLIDAY):
    workDay = []
    endDay = calendar.monthrange(year, month)[1]   #求本月天数
    if custom != []:
        workDay = [day for day in range(1, endDay + 1) if day not in custom]
    else:
        for day in range(1, endDay + 1):
            weekday = int(datetime.datetime(year, month, day).strftime("%w"))
            if weekday in range(1, 6):
                date = datetime.date(year, month, day)
                workDay.append(date)
    return workDay 

def judgeOneDay(curDayRecord):
    # flagAM1, flagAM2, flagPM1, flagPM2, flagTM = False, False, False, False, True
    flag = [False, False, False, False, True]
    timeInterval = [[MOR_BEGIN1, MOR_BEGIN2], [MOR_END1, MOR_END2], [AFT_BEGIN1, AFT_BEGIN2], [AFT_END1, AFT_END2]]
    for idx in range(4):
        start = datetime.datetime.strptime(timeInterval[idx][0], "%H:%M:%S").time()
        end = datetime.datetime.strptime(timeInterval[idx][1], "%H:%M:%S").time()
        flag[idx] = any([i for i in curDayRecord if i.time() > start and i.time() < end]) 

    totalSeconds = datetime.timedelta.total_seconds(curDayRecord[-1] - curDayRecord[0])
    if totalSeconds < TOTAL_TIME * 60:
        flag[4] = False
    return flag

def calc(dic, workDay):
    vioRecord = {}
    for name in dic:
        if not vioRecord.has_key(name):
            vioRecord[name] = []    
            vioRecord[name].append(0)
            vioRecord[name].append([])

        # 计算缺勤的天数
        attendDate = []
        for time in dic[name]:
            d = time.date()
            if d not in attendDate:
                attendDate.append(d)
        absentDate = [date for date in workDay if date not in attendDate]
        vioRecord[name][0] = len(absentDate)
        for date in absentDate:
            vioRecord[name][1].append(str(date) + "  AM1  AM2  PM1  PM2")

        # 计算迟到早退的天数
        curIndex = 0
        i = 0
        recList = dic[name]

        for date in workDay:
            curDayRecord = []
            while i < len(recList):
                if recList[i].date() == date:
                    curDayRecord.append(recList[i]) 
                    i += 1  
                elif recList[i].date() < date:
                    i += 1
                else:
                    break
            curIndex = i

            curDayRecord.sort()
            # print curDayRecord
            flags = judgeOneDay(curDayRecord)
            if not all(flags):
                vioRecord[name][0] += 1
                record = ""
                record += str(date)
                if not all(flags[0: 4]):
                    if flags[0] == False:
                        record += "  AM1"               # AM1 表示上午上班时间未打卡  
                    if flags[1] == False:   
                        record += "  AM2"               # AM2 表示上午下班时间未打卡 
                    if flags[2] == False:
                        record += "  PM1"               # PM1 表示下午上班时间未打卡 
                    if flags[3] == False:
                        record += "  PM2"               # PM2 表示下午下班时间未打卡 
                elif flags[4] == False: 
                    record += "  TM"                    # TM表示上班总时间不足
                vioRecord[name][1].append(record)    
            else:
                pass
    return vioRecord

def modify(filename, table):
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    sheetname = filename.split(".")[0]
    newFile = xlwt.Workbook() 
    newTable = newFile.add_sheet(sheetname, cell_overwrite_ok = True)
    content = []

    for rownum in range(1, nrows):
        row = table.row_values(rownum)
        if row:
            content.append(row)
            name = row[1]
            if name == ADD_NAME:
                for t in ADD_TIME:
                    time = row[3]
                    nextRow = table.row_values(rownum + 1)
                    if name == nextRow[1]:
                        nextTime = datetime.datetime.strptime(nextRow[3], "%Y-%m-%d %H:%M:%S")
                        time = datetime.datetime.strptime(time, "%Y-%m-%d %H:%M:%S")
                        t = datetime.datetime.strptime(t, "%Y-%m-%d %H:%M:%S")
                        if t > time and t < nextTime:
                            newRow = []
                            newRow = row[0: 3]
                            newRow.append(str(t))
                            newRow.extend(row[4: ncols])
                            content.append(newRow) 
                    
    newnrows = len(content)

    style = xlwt.XFStyle() 
    font = xlwt.Font()   #为样式创建字体
    font.name = 'Arial'
    font.bold = True
    style.font = font #为样式设置字体

    firstRow = table.row_values(0)
    for col in range(ncols):
        newTable.write(0, col, firstRow[col], style)

    # font.bold = False
    for row in range(newnrows):
        for col in range(ncols):
            newTable.write(row + 1, col, content[row][col])
                
    newFile.save(sheetname + "(new).xls")
    return

def output(vioRecord):
    for name in vioRecord:
        label = name + u": 违规 " + str(vioRecord[name][0]) + u" 次 " 
        print label.encode("gbk")
        for record in vioRecord[name][1]:
            print record 
        print
    return

def main(filename):
    lst, dic, table = excel_table_byindex(filename)
    year, month = filename.strip().split("-")[0:2]
    workDay = workday(int(year), int(month))
    vioRecord = calc(dic, workDay)
    output(vioRecord)
    # modify(filename, table)
    return 

if __name__=="__main__":
    # filename = '2014-12-wang.xls'
    filename = '2014-12-wang.xls'
    main(filename)
