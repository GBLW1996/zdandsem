import time
import pymysql
import xlrd
import os
import datetime
import xlwt
import pandas
import re


class CompareTable(object):
    def __init__(self, get_date_from, get_date_to):
        self.get_date_from = get_date_from
        self.get_date_to = get_date_to
        self.db = pymysql.connect(host='rr-bp1k48go5ks7gky746o.mysql.rds.aliyuncs.com', user='downer',
                                  password='baizhucc', database='downer', port=3306)
        self.cursor = self.db.cursor()
        self.desk_top = os.path.join(os.path.expanduser('~'), 'Desktop')

    def get_compare(self):
        sql = "SELECT top.*, ta.appname AS apper, SUM(top.hits) AS allhits FROM (SELECT * FROM storage_top WHERE " \
              "1 = 1 and `day` >='{}' and `day` <='{}') AS top LEFT JOIN tb_app AS ta ON ta.appid " \
              "= top.appid WHERE (ta.appname like '搜狗%' or ta.appname like '360%' or ta.appname like '百度%') and 1 " \
              "= 1 GROUP BY top.appid, top.sid ORDER BY top.`day` ASC, top.hits DESC".format(self.get_date_from, self.get_date_to)
        self.db.ping(reconnect=True)
        self.cursor.execute(sql)
        res_sem = self.cursor.fetchall()
        self.db.commit()
        rs_list = []  # 营销
        for i in res_sem:
            i = list(i)
            i[0] = '{}'.format(i[0])
            i[-1] = float(i[-1])
            i[1] = i[-2] + '-' + str(i[1])
            i.pop(4)
            i.pop(3)
            i.pop(-4)
            i.pop(-2)
            rs_list.append(i)
        sql1 = "SELECT top.*, ta.appname AS apper, SUM(top.hits) AS allhits FROM (SELECT * FROM storage_top WHERE " \
               "1 = 1 AND `day` >='{}' AND `day` <='{}') AS top LEFT JOIN tb_app AS ta ON ta.appid = " \
               "top.appid WHERE (LEFT(ta.appname,3)<>'360' AND LEFT(ta.appname,2)<>'搜狗' AND LEFT(ta.appname,2)<>'" \
               "百度') AND 1 = 1 GROUP BY top.appid, top.sid ORDER BY top.`day` ASC, top.hits DESC".format(self.get_date_from, self.get_date_to)
        self.db.ping(reconnect=True)
        self.cursor.execute(sql1)
        res_zd = self.cursor.fetchall()
        self.db.commit()
        rs_list1 = []  # 站点
        for i in res_zd:
            i = list(i)
            i[0] = '{}'.format(i[0])
            i[-1] = float(i[-1])
            i[1] = i[-2] + '-' + str(i[1])
            i.pop(4)
            i.pop(3)
            i.pop(-4)
            i.pop(-2)
            rs_list1.append(i)
        rs_list2 = []  # 站点点击高于100的
        for i in rs_list1:
            if i[-1] > 100.0:
                rs_list2.append(i)
        rs_list3 = []  # 词相似但点击小于100的
        for i in rs_list:
            for v in rs_list2:
                if ((v[3] in i[3] or i[3] in v[3]) and (i[-2].lower() == v[-2].lower())) and i[-1] < 100:
                # if (v[3] in i[3] or i[3] in v[3]) and i[-1] < 100:
                # if (((v[3] in i[3]) and (i[-2].lower() == v[-2].lower()) or (i[3] in v[3])) and (i[-2].lower() == v[-2].lower())) and i[-1] < 100:
                    rs_list3.append(i)
        rs_list4 = []  # 词相同或类似但点击小于100的去重后
        for item in rs_list3:
            if not item in rs_list4:
                rs_list4.append(item)
        zd_removal = []  # 站点大于100点击的所有词
        for item in rs_list2:
            zd_removal.append(item[3])
        zd_all = []  # 站点大于100点击的所有词去重后
        for i in zd_removal:
            if not i in zd_all:
                zd_all.append(i)
        # print(zd_all)
        # print(len(zd_all))
        sem_all = []  # 所有的营销的词
        for i in rs_list:
            sem_all.append(i[3])
        zd_have_sem_no = []  # 站点有，营销没有的词
        for i in zd_all:
            if i not in sem_all:
                zd_have_sem_no.append(i)
        # print(zd_have_sem_no)
        # print(len(zd_have_sem_no))
        zd_h_msg = []  # 站点有，营销没有的词所对应的所有站点详细信息
        for i in rs_list2:
            for v in zd_have_sem_no:
                if v in i:
                    zd_h_msg.append(i)
        rs_list5 = []  # 站点有，营销没有的词所对应的所有站点详细信息去重
        for item in zd_h_msg:
            if not item in rs_list5:
                rs_list5.append(item)
        # print(rs_list5)
        return rs_list, rs_list1, rs_list4, rs_list5

    def write_excel(self):
        sem_res, zd_res, cp_res, zd_h_res = self.get_compare()
        self.write_func(sem_res, '营销.xls', '营销')
        # self.write_func(zd_res, '站点.xls', '站点')
        self.write_func(cp_res, '类似词营销量级低于100.xls', '营销类似词且数量小于100')
        self.write_func(zd_h_res, '站点100以上营销没有的词.xls', '站点100以上营销没有的词')

    def write_func(self, exc_list, book_name, sheet_name):
        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet(u'{}'.format(sheet_name), cell_overwrite_ok=True)
        list_title = ['日期', '渠道号', '资源id', '资源名称', '扩展名', '次数']
        for i in list_title:
            worksheet.write(0, list_title.index(i), i)
        i = 0
        for data in exc_list:
            for j in range(len(data)):
                worksheet.write(i + 1, j, data[j])
            i = i + 1
        """ 
        设置单元格高度
        worksheet.row(0).height_mismatch = True
        worksheet.row(0).height = 20 * 30
        """
        worksheet.col(0).width = 256 * 16  # 一字节宽度*字节数
        worksheet.col(1).width = 256 * 25
        worksheet.col(3).width = 256 * 40
        workbook.save(self.desk_top + '\{}'.format(book_name))
        # dataframe_zd = pandas.DataFrame(self.get_compare())  # 用pandas写
        # dataframe_zd.to_excel(self.desk_top + '\站点.xls', sheet_name='站点表')


if __name__ == "__main__":
    # get_date_from = input('请输入要同步的开始日期，格式参考：1996-09-18\n')  # 起始日期
    # get_date_to = input('请输入要同步的结束日期，格式参考：1996-09-18\n')  # 结束日期
    # print('同步中，稍等...')
    get_date_from = ''
    get_date_to = ''

    def judge_date(get_date):
        result1 = re.findall('^\d{4}-\d{2}-\d{2}$', get_date)
        try:
            date_p = datetime.datetime.strptime(get_date, '%Y-%m-%d').date()
        except:
            print('请输入正确日期')
            return False
        yesterday = datetime.date.today() - datetime.timedelta(days=1)
        if int(get_date[0:4]) < 2017:
            print('最早输入2017年日期')
            return False
        elif not result1:
            print('日期格式错误')
            return False
        elif time.mktime(time.strptime(get_date, '%Y-%m-%d')) > time.mktime(time.strptime(str(yesterday), '%Y-%m-%d')):
            print('此日期还未写入数据')
            return False
        else:
            return True
    while True:
        get_date_from = input('请输入要同步的开始日期，格式参考：1996-09-18\n')
        if judge_date(get_date_from):
            get_date_to = input('请输入要同步的结束日期，格式参考：1996-09-18\n')
        else:
            continue
        if judge_date(get_date_to):
            if time.mktime(time.strptime(get_date_from, '%Y-%m-%d')) > time.mktime(time.strptime(get_date_to, '%Y-%m-%d')):
                print('输入的开始日期大于结束日期')
                continue
            else:
                print('同步中，请稍等')
                break
        else:
            continue

    cp = CompareTable(get_date_from, get_date_to)
    cp.write_excel()
    print('抓取完成，去你的桌面上找“营销.xls”、“类似词营销量级低于100.xls”、“站点100以上营销没有的词.xls”')
    time.sleep(3)