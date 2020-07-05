from PyQt5.QtWidgets import QWidget, QApplication, QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView, QComboBox, QLabel
from PyQt5.QtCore import QDate, QTime, QStringListModel
from PyQt5.QtGui import QIntValidator
from  PyQt5.QtGui import QColor

# from PyQt5 import QtCore
# import numpy as np
import pandas as pd
from ui_easy1 import *
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
# from lxml import etree
import time
from selenium.webdriver.common.action_chains import ActionChains
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException


BTN_Save = "//*[@id='ctl00xcphToolbarxuwtToolbar_Item_4_img']"
BTN_Search = "//*[@id='ctl00xcphToolbarxuwtToolbar_Item_5_img']"
BTN_New = "//*[@id='ctl00xcphToolbarxuwtToolbar_Item_19_img']"
BTN_Submit1 = "//*[@id='MC_TC__ctl2_ctl00__964']"
BTN_Submit2 = "//*[@id='MC_TC__ctl2_ctl00__5025']"
BTN_ACK = "//*[@id='MC_TC__ctl3_ctl00_Splitter_10007_4_1_2__1985']"
USERNAME = '//*[@name="tbUser"]'
PASSWORD = '//*[@name="tbPassword"]'
LOGIN = '//*[@name="btnLogin"]'
GD = {
    '查询': {
        'TAB_btn': "//td[@tabid='MC_TC,0']",
        '编号': "//*[@id='MC_TC__ctl0_ctl00__5002']",
        '年份': {
            'combo_btn': "//*[@id='MC_TC__ctl0_ctl00__102']/div/table/tbody/tr/td[2]/img",
            '2017': "//*[@data-ig='x:898879528.26:adr:18']",
            '2018': "//*[@data-ig='x:898879528.27:adr:19']",
            '2019': "//*[@data-ig='x:898879528.28:adr:20']",
            '2020': "//*[@data-ig='x:898879528.29:adr:21']",
        },
    },
    '工单': {
        'TAB_btn': "//td[@tabid='MC_TC,2']",
        '优先级': {
            'combo_btn': '//*[@id="MC_TC__ctl2_ctl00__314"]/div/table/tbody/tr/td[2]/img',
            'input_adr': "//*[@data-ig='x:1317819027.2:mkr:Input']",
            '立刻': "//*[@data-ig='x:1317819027.8:adr:0']",
            '24小时内处理': "//*[@data-ig='x:1317819027.9:adr:1']",
            '48小时内处理': "//*[@data-ig='x:1317819027.10:adr:2']",
            '72小时内处理': "//*[@data-ig='x:1317819027.11:adr:3']",
            '168小时内处理': "//*[@data-ig='x:1317819027.12:adr:4']",
            '停机处理': "//*[@data-ig='x:1317819027.13:adr:5']",
        },
        '维修类型': {
            'combo_btn': '//*[@id="MC_TC__ctl2_ctl00__311"]/div/table/tbody/tr/td[2]/img',
            'input_adr': "//*[@data-ig='x:126177939.2:mkr:Input']",
            '定期工作': "//*[@data-ig='x:126177939.8:adr:0']",
            '改进性检修': "//*[@data-ig='x:126177939.9:adr:1']",
            '故障检修': "//*[@data-ig='x:126177939.10:adr:2']",
            '状态检修': "//*[@data-ig='x:126177939.11:adr:3']",
        },
        '机组': {
            'combo_btn': '//*[@id="MC_TC__ctl2_ctl00__3002"]/div/table/tbody/tr/td[2]/img',
            'input_adr': "//*[@data-ig='x:508519679.2:mkr:Input']",
            '当前值': '',
            '目标值': '',
            '丰泰#1机组': "//*[@data-ig='x:508519679.8:adr:0']",
            '丰泰#2机组': "//*[@data-ig='x:508519679.9:adr:1']",
            '丰泰公用': "//*[@data-ig='x:508519679.10:adr:2']",
            '科林新#3机组': "//*[@data-ig='x:508519679.11:adr:3']",
            '科林新#4机组': "//*[@data-ig='x:508519679.12:adr:4']",
            '科林公用': "//*[@data-ig='x:508519679.13:adr:5']",
        },
        '专业': {
            'combo_btn': '//*[@id="MC_TC__ctl2_ctl00__717"]/div/table/tbody/tr/td[2]/img',
            'input_adr': "//*[@data-ig='x:256421229.2:mkr:Input']",
            '汽机专业': "//*[@data-ig='x:256421229.8:adr:0']",
            '锅炉专业': "//*[@data-ig='x:256421229.9:adr:1']",
            '电气专业': "//*[@data-ig='x:256421229.10:adr:2']",
            '热控专业': "//*[@data-ig='x:256421229.11:adr:3']",
            '输煤专业': "//*[@data-ig='x:256421229.12:adr:4']",
            '脱硫专业': "//*[@data-ig='x:256421229.13:adr:5']",
            '化学专业': "//*[@data-ig='x:256421229.14:adr:6']",
            '土建专业': "//*[@data-ig='x:256421229.15:adr:7']",
            '远动专业': "//*[@data-ig='x:256421229.16:adr:8']",
        },
        '机组类别': {
            'combo_btn': '//*[@id="MC_TC__ctl2_ctl00__3004"]/div/table/tbody/tr/td[2]/img',
            'input_adr': "//*[@data-ig='x:508519685.2:mkr:Input']",
            '内蒙古丰泰发电厂': "//*[@data-ig='x:508519685.8:adr:0']",
            '呼和浩特科林热电厂': "//*[@data-ig='x:508519685.9:adr:1']",
            '科林城发热力': "//*[@data-ig='x:508519685.10:adr:2']",
        },
        '负责人': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__990']", '当前值': '', '目标值': '', },
        '工作内容': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__348']", '当前值': '', '目标值': '', },
        '开始日期': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__972']", '当前值': '', '目标值': '', },
        '开始时间': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__872']", '当前值': '', '目标值': '', },
        '结束日期': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__973']", '当前值': '', '目标值': '', },
        '结束时间': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__873']", '当前值': '', '目标值': '', },
        '电厂编码': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__960']", '当前值': '', '目标值': '', },

    },
    '工单分项': {
        'TAB_btn': "//td[@tabid='MC_TC,3']",
        '分项表格': "//*[@data-ig = 'x:914667624.17:adr:0:tag:']",
        '班组': "//*[@id='MC_TC__ctl3_ctl00_Splitter_10007_4_1_3__960']",
    },
    '连接': {
        'TAB_btn': "//td[@tabid='MC_TC,4']",
        '工作票': "//*[@id='ctl00MCTCctl4ctl00Splitter100071111ctl00uwLinkTree_1_2_1_6']/span[3]",
        '打开工作票': "//*[@id='MC_TC__ctl4_ctl00_Splitter_10007_11_1_1_ctl01_uwPopMenu1n3']/td/table",
        '创建工作票': "//*[@id='MC_TC__ctl4_ctl00_Splitter_10007_11_1_1_ctl01_uwPopMenu1n0']/td/table/tbody/tr/td/a/div",
    },

}
GZP = {
    '工作票': {
        'TAB_btn': "//td[@tabid='MC_TC,2']",
        '工作票类型': {
            'combo_btn': "//*[@id='MC_TC__ctl2_ctl00__107']/div/table/tbody/tr/td[2]/img",
            'input_adr': '//*[@data-ig="x:1822111954.2:mkr:Input"]',
            '电气工作票': "//*[@data-ig='x:1822111954.8:adr:0']",
            '热力机械工作票': "//*[@data-ig='x:1822111954.9:adr:1']",
            '外包电气工作票': "//*[@data-ig='x:1822111954.10:adr:2']",
            '外包热力机械工作票': "//*[@data-ig='x:1822111954.11:adr:3']",
        },
        '专业': {
            'combo_btn': "//*[@id='MC_TC__ctl2_ctl00__252']/div/table/tbody/tr/td[2]/img",
            'input_adr': '//*[@data-ig="x:495527167.2:mkr:Input"]',
            '汽机专业': "//*[@data-ig='x:495527167.8:adr:0']",
            '锅炉专业': "//*[@data-ig='x:495527167.9:adr:1']",
            '电气专业': "//*[@data-ig='x:495527167.10:adr:2']",
            '热控专业': "//*[@data-ig='x:495527167.11:adr:3']",
            '输煤专业': "//*[@data-ig='x:495527167.12:adr:4']",
            '脱硫专业': "//*[@data-ig='x:495527167.13:adr:5']",
            '化学专业': "//*[@data-ig='x:495527167.14:adr:6']",
            '土建专业': "//*[@data-ig='x:495527167.15:adr:7']",
            '远动专业': "//*[@data-ig='x:495527167.16:adr:8']",
        },
        '机组': {
            'combo_btn': "//*[@id='MC_TC__ctl2_ctl00__5043']/div/table/tbody/tr/td[2]/img",
            'input_adr': '//*[@data-ig="x:260479746.2:mkr:Input"]',
            '丰泰#1机组': "//*[@data-ig='x:260479746.8:adr:0']",
            '丰泰#2机组': "//*[@data-ig='x:260479746.9:adr:1']",
            '丰泰公用': "//*[@data-ig='x:260479746.10:adr:2']",
            '科林新#3机组': "//*[@data-ig='x:260479746.11:adr:3']",
            '科林新#4机组': "//*[@data-ig='x:260479746.12:adr:4']",
            '科林公用': "//*[@data-ig='x:260479746.13:adr:5']",
        },
        '总人数': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__156']", '当前值': '', '目标值': '', },
        '负责人': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__112']", '当前值': '', '目标值': '', },
        '班组成员': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__153']", '当前值': '', '目标值': '', },
        '工作内容及地点': {'input_adr': "//*[@id='MC_TC__ctl2_ctl00__120']", '当前值': '', '目标值': '', },

    },
    '电厂编码': {
        'TAB_btn': "//td[@tabid='MC_TC,3']",
        '隔离': {
            'combo_btn': "//*[@data-ig='x:683059316.4:mkr:ButtonImage']",
            'input_adr': '',
            '电气隔离': "//*[@data-ig='x:683059316.8:adr:0']",
            '机械隔离': "//*[@data-ig='x:683059316.9:adr:1']",
            '电气及机械隔离': "//*[@data-ig='x:683059316.10:adr:2']",
        },
        '编码表格': "//*[@data-ig = 'x:1317427084.17:adr:0:tag:']",
        'radio_btn_电气': "//*[@id='MC_TC__ctl3_ctl00__247_img']",
        'radio_btn_机械': "//*[@id='MC_TC__ctl3_ctl00__248_img']",
        'radio_value_电气': "//*[@id='MC_TC__ctl3_ctl00__247']",
        'radio_value_机械': "//*[@id='MC_TC__ctl3_ctl00__248']",

    },
    '隔离措施': {
        'TAB_btn': "//td[@tabid='MC_TC,5']",
        '隔离状态': {
            'combo_btn': "//*[@id='MC_TC__ctl5_ctl00__231']/div/table/tbody/tr/td[2]/img",
            'input_adr': '//*[@data-ig="x:1289034592.2:mkr:Input"]',
            '备用': "//*[@data-ig='x:1289034592.8:adr:0']",
            '不停电': "//*[@data-ig='x:1289034592.9:adr:1']",
            '断开': "//*[@data-ig='x:1289034592.10:adr:2']",
            '关': "//*[@data-ig='x:1289034592.11:adr:3']",
            '关闭、停电': "//*[@data-ig='x:1289034592.12:adr:4']",
            '合上': "//*[@data-ig='x:1289034592.13:adr:5']",
            '解除自动': "//*[@data-ig='x:1289034592.14:adr:6']",
            '开': "//*[@data-ig='x:1289034592.15:adr:7']",
            '停电': "//*[@data-ig='x:1289034592.16:adr:8']",
            '停运': "//*[@data-ig='x:1289034592.17:adr:9']",
            '停运断电': "//*[@data-ig='x:1289034592.18:adr:10']",
            '退出保护': "//*[@data-ig='x:1289034592.19:adr:11']",
            '退出监视': "//*[@data-ig='x:1289034592.20:adr:12']",
            '退出联锁': "//*[@data-ig='x:1289034592.21:adr:13']",
        },
        '安全标示': {
            'combo_btn': "//*[@id='MC_TC__ctl5_ctl00__234']/div/table/tbody/tr/td[2]/img",
            'input_adr': '//*[@data-ig="x:1814291616.2:mkr:Input"]',
            '禁止操作，有人工作': "//*[@data-ig='x:1814291616.8:adr:0']",
            '禁止合闸，有人工作': "//*[@data-ig='x:1814291616.9:adr:1']",
            '禁止攀登': "//*[@data-ig='x:1814291616.10:adr:2']",
            '由此上下': "//*[@data-ig='x:1814291616.11:adr:3']",
            '在此工作': "//*[@data-ig='x:1814291616.12:adr:4']",
            '止步，高压危险': "//*[@data-ig='x:1814291616.13:adr:5']",
        },
        '隔离切换操作': {
            'combo_btn': "//*[@id='MC_TC__ctl5_ctl00__216']/div/table/tbody/tr/td[2]/img",
            'input_adr': '',
            '记录陈述': "//*[@data-ig='x:1700213290.8:adr:0']",
            '隔离切换操作': "//*[@data-ig='x:1700213290.9:adr:1']",
        },
        '措施类别': {
            'combo_btn': "//*[@data-ig='x:2074329851.4:mkr:ButtonImage']",
            'input_adr': '',
            '热机票-1': "//*[@data-ig='x:2074329851.8:adr:0']",
            '热机票-2': "//*[@data-ig='x:2074329851.9:adr:1']",
        },
        '电厂编码': "//*[@id='MC_TC__ctl5_ctl00__960']",
        '措施内容': "//*[@id='MC_TC__ctl5_ctl00__218']",
        '措施表格': "//*[@data-ig='x:1317417514.18:mkr:rows:nw:1']",

    },
    '危险预控': {
        'TAB_btn': "//td[@tabid='MC_TC,6']",
        '危险辨识': "//*[@id='MC_TC__ctl6_ctl00__371']",
        '预控措施': "//*[@id='MC_TC__ctl6_ctl00__5002']",
        '预控表格': "//*[@data-ig = 'x:1317421711.16:mkr:rows:nw:1']"
    },
}
EMPLOYEE_INFO={
    '胡毅' :      ['80180258', '80180258'],
    '周军' :      ['80180266', '80180266'],
    '王志英' :    ['80181079', 'AC2564'],
    '梁春红' :    ['80182666', '80182666'],
    '云雯' :      ['82026854', '1'],
    '车璐' :      ['82043074', '1'],
    '范舒婷' :    ['82043072', '4'],

}
FILEPATH = r'data\easy.xlsx'
PATH = r"d:\app\chromedriver.exe"
AIMDICT = {}


class Window(QWidget, Ui_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        #初始化数据
        self.first_dataframe = None
        self.current_dataframe = None
        self.current_dataframe_dingqi = None
        self.current_dataframe_weixiu = None

        self.gelidanhao = None
        self.bool_login=False
        self.current_year = ''
        self.selected_index = None
        self.current_year = '模板'
        self.is_edit = False
        self.select_xuhao = None
        self.xuhao_last = None

        self.get_df_from_excel(FILEPATH)
        self.main_table.clicked.connect(self.main_table_clicked)
        self.main_table.setSelectionBehavior(QTableWidget.SelectRows)       #整行选择
        self.main_table.setEditTriggers(QTableWidget.NoEditTriggers)        #无法编辑
        self.main_table.setSelectionMode(QTableWidget.SingleSelection)      #只能选择一行
        self.glcs_table.setSelectionBehavior(QTableWidget.SelectRows)        #整行选择
        self.ykcs_table.setSelectionBehavior(QTableWidget.SelectRows)       #整行选择
        # self.year_combo.activated.connect(self.year_change)
        self.search_btn.clicked.connect(lambda: self.search_btn_clicked())
        self.filter_text.returnPressed.connect(self.search_btn_clicked)
        self.start_btn.clicked.connect(self.start_autowork_clicked)
        self.btn_glcs_addrow.clicked.connect(self.add_rows)
        self.btn_glcs_delrow.clicked.connect(self.remove_rows)
        self.find_btn.clicked.connect(self.find_gd_clicked)
        self.denglu.clicked.connect(self.login)
        # self.checkBox.clicked.connect(self.check_dingqi)
        self.checkBox.stateChanged.connect(self.check_dingqi)
        self.find_btn.setEnabled(False)

        # self.start_btn.setEnabled(True)


        self.save_btn.clicked.connect(lambda :self.save_dataframe(self.current_dataframe))
        self.del_btn.clicked.connect(self.del_dataframe)
        self.btn_replace.clicked.connect(self.glcs_replace)
        # self.text_find.textChanged.connect(self.text_find_changed)
        intvalidate = QIntValidator(self)
        intvalidate.setRange(1, 8888)
        self.lineEdit.setValidator(intvalidate)
        self.begindate.setDate(QDate.currentDate())
        self.begintime.setTime(QTime.currentTime())
        etime = QTime.fromString('18:00', 'hh:mm')
        if QTime.currentTime() > etime:
            # print('超过，加1')
            self.enddate.setDate(QDate.currentDate().addDays(1))
        else:
            self.enddate.setDate(QDate.currentDate())
        self.endtime.setTime(QTime.fromString('18:00', 'hh:mm'))
        self.Auto = None

        # self.save_btn.setEnabled(False)
        self.del_btn.setEnabled(True)
        #========编辑状态判断============编辑状态判断===============编辑状态判断===========编辑状态判断====================

        self.youxianji.activated.connect(self.edit_state)
        self.weixiuleixing.activated.connect(self.edit_state)
        self.zhuanye.activated.connect(self.edit_state)
        self.banzu.activated.connect(self.edit_state)
        self.jizuleibie.activated.connect(self.edit_state)
        self.jizu.activated.connect(self.edit_state)
        self.gzpleixing.activated.connect(self.edit_state)
        self.gelileixing.activated.connect(self.edit_state)
        self.gelizhuangtai.activated.connect(self.edit_state)
        self.anquanbiaoshi.activated.connect(self.edit_state)
        self.mvp.textEdited.connect(self.edit_state)
        self.chengyuan.textEdited.connect(self.edit_state)
        self.bianma.textEdited.connect(self.edit_state)
        self.sum.textEdited.connect(self.edit_state)
        self.glcs_table.itemDoubleClicked.connect(self.edit_state)
        self.ykcs_table.itemDoubleClicked.connect(self.edit_state)
        #========编辑状态判断============编辑状态判断===============编辑状态判断===========编辑状态判断====================

    def callback_auto_write(self,index,msg):
        if index == 1:
            pass
        elif index == 2:
            self.display.setText(msg)
        elif index == 3:
            if (msg == 'END'):
                self.start_btn.setEnabled(True)
            pass

    #查找工单后返回的字典，转换为CURRENT DATAFRAME,新加序号=self.xuhao_last + 1
    def callback_search_dict(self,index,dict):
        if index == 1:
            pass
        elif index == 2:
            pass
        elif index == 3:
            mylist = []
            mylist.append(dict)
            df = pd.DataFrame(mylist)
            df['序号'] = self.xuhao_last + 1   #new add  xuhao
            self.current_dataframe = df
            self.insert_data(self.current_dataframe)

    def callback_search(self,index,msg):
        if index ==1 :
            self.progressBar.setValue(int(msg))
        elif index ==2 :
            self.display.setText(msg)
        elif index ==3:
            if (msg == 'END'):
                self.find_btn.setEnabled(True)

        
    def callback_login(self,index,msg):
        if index == 1:
            self.progressBar.setValue(int(msg))
        elif index == 2:
            self.display.setText(msg)
            if msg == '登陆成功':
                print('success')
                self.find_btn.setEnabled(True)
                self.start_btn.setEnabled(True)


    def login(self):
        self.Auto=AutoOperator(self)
        tempstr = self.fuzeren.currentText()
        self.display.setText('正在登录...')
        self.thread1 = Login_Thread(self.Auto,EMPLOYEE_INFO[tempstr])
        self.thread1._login_sign.connect(self.callback_login)
        self.thread1.start()


    #定期工作切换
    def check_dingqi(self):
        self.insert_data(self.current_dataframe)
        self.clear_info()


    def edit_state(self):
        self.is_edit = True
        print('正在编辑！')


    #搜索按钮点击事件
    def search_btn_clicked(self):
        filtertext = self.filter_text.text()   #搜索的关键字，以空格键隔开
        display_dataframe = self.filter_dateframe(self.current_dataframe, 'search', filtertext)


        self.insert_data(display_dataframe)
        # self.save_btn.setEnabled(False)

    # dateframe数据过滤
    def filter_dateframe(self, dateframe, type, filter_str = ''):
        if type == "search":
            if filter_str == '':
                return dateframe
            temp_list = re.split(r' +', filter_str)
            #去除列表中空字符
            def no_empty(s):
                return s and s.strip()
            filter_list = list(filter(no_empty, temp_list))  # 过滤字符列表filter_list
            del_list = []
            for  row in dateframe.itertuples():
                for text in filter_list:
                    #关键字只搜索【班组成员】和【工作内容】
                    if str(row.班组成员).find(text) > -1 or str(row.工作内容).find(text) > -1:
                        pass
                    else:
                        del_list.append(row[0])  # 需要删除的索引号，加入列表
                        break
            temp_df = dateframe.drop(labels=del_list)  # 按列表中索引删除后，索引保持，需要重建索引,inplace=True表示原DATAFRAME改变
            # temp_df = temp_df.reset_index(drop=True)  # 重建索引
            return temp_df
        elif type == 'dingqi-fenli':
            del_list = [] #定期列表
            del_list2 = [] #维修列表
            for item in dateframe.itertuples():
                if str(item.维修类型).find('定期工作') > -1:
                    del_list.append(item[0])
                else:
                    del_list2.append(item[0])
            weixiu_df = dateframe.drop(labels=del_list).drop_duplicates(subset='序号', keep= 'first')   #去掉定期后的data
            dingqi_df = dateframe.drop(labels=del_list2).drop_duplicates(subset='序号', keep= 'first')    #去掉维修后的data
            return dingqi_df ,weixiu_df


    def save_dataframe(self,gd_num):
        #判断first_datafram中序号 是否相同
        # if int(self.select_xuhao) in self.first_dataframe['序号'].values:
        #     print('exist,不保存！')
        #     return

        # 用【负责人】判断 是否选中ITEM
        # print(self.fuzeren.currentText())
        if self.mvp.text() == '':
            print('没有选中ITEM')
        else:
            aim_dict = {}
            self.info_to_dict(aim_dict)
            aim_dict['序号'] = self.xuhao_last + 1
            # print(aim_dict)
            mylist = []
            mylist.append(aim_dict)
            df = pd.DataFrame(mylist)
            print(df)

            #1、读入当前文件SHEET为模板dataframe
            df_moban = pd.read_excel(FILEPATH, '模板')
            #2、将current dataframe添加序号列后，追加到模板dataframe
            df_moban = df_moban.append(df, ignore_index= True)
            self.current_dataframe = df_moban
            self.insert_data(self.current_dataframe)
            writer = pd.ExcelWriter(FILEPATH)
            df_moban.to_excel(excel_writer=writer, sheet_name='模板', index= False)
            writer.save()
            writer.close()
            self.clear_info()



    def del_dataframe(self):  #删除选中data
        if self.mvp.text() == '':
            print('没有选中ITEM')
            return
        else:

            del_list = []
            # df = self.current_dataframe[self.current_dataframe['序号'].isin([self.select_xuhao])]
            # 遍历df找到序号对应的索引
            for row in self.current_dataframe.itertuples():
                if self.select_xuhao == str(row[1]).strip():
                    del_list.append(row[0])
            self.current_dataframe.drop(labels=del_list, inplace=True)
            writer = pd.ExcelWriter(FILEPATH)
            self.current_dataframe.to_excel(excel_writer=writer, sheet_name='模板', index=False)
            writer.save()
            writer.close()
            self.insert_data(self.current_dataframe)
            self.clear_info()


    #通过年份，获取主dateframe数据，并插入数据到主TABLE
    def get_df_from_excel(self, file_path, years='模板'):
        self.first_dataframe = pd.read_excel(file_path)
        self.current_dataframe = self.first_dataframe
        #获取最后一行序号
        self.insert_data(self.current_dataframe)
        # self.xuhao_last = self.first_dataframe.tail(1)['序号'].values[0]
        # print('最后序号为{}'.format(self.xuhao_last))



    #隔离措施TABLE增加行
    def add_rows(self):
        currentrow = self.glcs_table.currentRow()
        rowcount = self.glcs_table.rowCount()
        if currentrow == -1:
            self.glcs_table.insertRow(rowcount)
        else:
            self.glcs_table.insertRow(currentrow+1)

    # 隔离措施TABLE删除行
    def remove_rows(self):
        # currentrow = self.glcs_table.currentRow()
        selections = self.glcs_table.selectionModel()
        selectionslist = selections.selectedRows()
        rows = []
        for r in selectionslist:
            rows.append(r.row())
        rows.reverse()
        for i in rows:
            self.glcs_table.removeRow(i)


    #插入数据到主TABLE
    def insert_data(self, current_df):
        self.xuhao_last = current_df.tail(1)['序号'].values[0]
        print('最后序号为{}'.format(self.xuhao_last))
        dingqi, weixiu = self.filter_dateframe(current_df, type='dingqi-fenli')
        if self.checkBox.isChecked():
            temp_df = dingqi
        else:
            temp_df = weixiu
        selected_df = temp_df[['序号', '工单号', '负责人', '机组', '工作内容', '维修类型', '隔离单号']]
        # print(selected_df)
        rows = selected_df.shape[0]
        columns = selected_df.shape[1]
        headers = selected_df.columns.values.tolist()
        self.main_table.clear()
        self.main_table.setRowCount(rows)
        self.main_table.setColumnCount(columns)
        self.main_table.setHorizontalHeaderLabels(headers)
        self.main_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.main_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.main_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.main_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.main_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.main_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        for i in range(rows):
            for j in range(columns):
                newitem = QTableWidgetItem(str(selected_df.iat[i, j]))
                # if j == 1:
                #     newitem = QTableWidgetItem(str(selected_df.iat[i, j]).split('.')[0])
                self.main_table.setItem(i, j, newitem)



    #主TABLE点击事件
    def main_table_clicked(self, item):
        if self.is_edit == True:
            if self.select_xuhao != None:
                res = QMessageBox.information(None, 'title', '信息已改变，是否离开', QMessageBox.Yes | QMessageBox.No)
                if res == QMessageBox.No:
                    self.main_table.selectRow(self.selected_index)
                    return
            self.is_edit = False

        self.selected_index = item.row()            #选中行的行数
        # current_row = item.row()
        currnet_df = self.current_dataframe
        self.select_xuhao = self.main_table.item(self.selected_index, 0).text()         #获得选中行第0列，即序号
        #通过序号获取DATA
        df = currnet_df[currnet_df['序号'].isin([self.select_xuhao])]
        def f(x):
            return '无' if x.isnull().any() else x.item()

        self.gelidanhao = str(f(df['隔离单号']))
        self.youxianji.setCurrentText(f(df['优先级']))
        print(f(df['维修类型']))
        self.weixiuleixing.setCurrentText(f(df['维修类型']))
        self.mvp.setText(f(df['负责人']))
        self.jizu.setCurrentText(f(df['机组']))
        self.zhuanye.setCurrentText(f(df['专业']))
        self.bianma.setText(f(df['电厂编码']))
        self.content.setText(f(df['工作内容']))
        self.jizuleibie.setCurrentText(f(df['机组类别']))
        self.banzu.setCurrentText(f(df['班组']))
        self.gzpleixing.setCurrentText(f(df['工作票类型']))
        self.sum.setText(str(f(df['总人数'])).split('.')[0])
        self.chengyuan.setText(f(df['班组成员']))
        self.contentaddr.setText(f(df['工作内容及地点']))
        self.gelileixing.setCurrentText(f(df['隔离类型']))
        self.gelizhuangtai.setCurrentText(f(df['隔离状态']))
        self.anquanbiaoshi.setCurrentText(f(df['安全标示']))
        glcs = f(df['隔离措施'])
        ykcs = f(df['预控措施'])
        self.insert_glcs_data(glcs)
        self.insert_ykcs_data(ykcs)


    def info_to_dict(self, out_dict):
        str_glcs = self.get_glcs_data()
        str_ykcs = self.get_ykcs_data()
        # out_dict = {}
        out_dict['序号'] =self.select_xuhao
        out_dict['工单号'] = self.select_xuhao
        out_dict['隔离单号'] = self.gelidanhao
        out_dict['优先级'] = self.youxianji.currentText()
        out_dict['维修类型'] = self.weixiuleixing.currentText()
        out_dict['负责人'] = self.mvp.text()
        out_dict['机组'] = self.jizu.currentText()
        out_dict['专业'] = self.zhuanye.currentText()
        out_dict['电厂编码'] = self.bianma.text()
        out_dict['工作内容'] = self.content.text()
        out_dict['机组类别'] = self.jizuleibie.currentText()
        out_dict['开始日期'] = QDate.toString(self.begindate.date(), 'yyyy/MM/dd')
        out_dict['开始时间'] = QTime.toString(self.begintime.time(), 'hh:mm')
        out_dict['结束日期'] = QDate.toString(self.enddate.date(), 'yyyy/MM/dd')
        out_dict['结束时间'] = QTime.toString(self.endtime.time(), 'hh:mm')
        out_dict['班组'] = self.banzu.currentText()
        out_dict['工作票类型'] = self.gzpleixing.currentText()
        out_dict['总人数'] = self.sum.text()
        out_dict['班组成员'] = self.chengyuan.text()
        out_dict['工作内容及地点'] = self.contentaddr.text()
        out_dict['隔离类型'] = self.gelileixing.currentText()
        out_dict['隔离状态'] = self.gelizhuangtai.currentText()
        out_dict['安全标示'] = self.anquanbiaoshi.currentText()
        out_dict['隔离措施'] = str_glcs
        out_dict['预控措施'] = str_ykcs
        return  out_dict

        pass
    #点击开始按钮，开始自动办理工作票
    def start_autowork_clicked(self):
        try:
            str_glcs = self.get_glcs_data()
            str_ykcs = self.get_ykcs_data()
            AIMDICT = {}
            self.info_to_dict(AIMDICT)
            # AIMDICT['工单号'] = self.select_xuhao
            # AIMDICT['隔离单号'] = self.gelidanhao
            # AIMDICT['优先级'] = self.youxianji.currentText()
            # AIMDICT['维修类型'] = self.weixiuleixing.currentText()
            # AIMDICT['负责人'] = self.mvp.text()
            # AIMDICT['机组'] = self.jizu.currentText()
            # AIMDICT['专业'] = self.zhuanye.currentText()
            # AIMDICT['电厂编码'] = self.bianma.text()
            # AIMDICT['工作内容'] = self.content.text()
            # AIMDICT['机组类别'] = self.jizuleibie.currentText()
            # AIMDICT['开始日期'] = QDate.toString(self.begindate.date(), 'yyyy/MM/dd')
            # AIMDICT['开始时间'] = QTime.toString(self.begintime.time(), 'hh:mm')
            # AIMDICT['结束日期'] = QDate.toString(self.enddate.date(), 'yyyy/MM/dd')
            # AIMDICT['结束时间'] = QTime.toString(self.endtime.time(), 'hh:mm')
            # AIMDICT['班组'] = self.banzu.currentText()
            # AIMDICT['工作票类型'] = self.gzpleixing.currentText()
            # AIMDICT['总人数'] = self.sum.text()
            # AIMDICT['班组成员'] = self.chengyuan.text()
            # AIMDICT['工作内容及地点'] = self.contentaddr.text()
            # AIMDICT['隔离类型'] = self.gelileixing.currentText()
            # AIMDICT['隔离状态'] = self.gelizhuangtai.currentText()
            # AIMDICT['安全标识'] = self.anquanbiaoshi.currentText()
            # AIMDICT['隔离措施'] = str_glcs
            # AIMDICT['危险预控'] = str_ykcs

            if AIMDICT['优先级'] == '无': self.display.text('优先级==NULL') ; return
            if AIMDICT['维修类型'] == '无': self.display.text('维修类型==NULL') ; return
            if AIMDICT['负责人'] == '无': self.display.text('负责人==NULL') ; return
            if AIMDICT['机组'] == '无': self.display.text('机组==NULL') ; return
            if AIMDICT['专业'] == '无': self.display.text('专业==NULL') ; return
            if AIMDICT['电厂编码'] == '无': self.display.text('电厂编码==NULL') ; return
            if AIMDICT['工作内容'] == '无': self.display.text('工作内容==NULL') ; return
            if AIMDICT['机组类别'] == '无': self.display.text('机组类别==NULL') ; return
            if AIMDICT['维修类型'] != '定期工作' and AIMDICT['工作票类型'] == '无': self.display.text('工作票类型==NULL') ; return

            # print(AIMDICT)
            # return
            self.start_btn.setEnabled(False)
            # -----------------TEST--------------------

            # -----------------TEST--------------------
            self.thread3 = Auto_Write_Thread(self.Auto,AIMDICT)
            self.thread3._auto_write_sign.connect(self.callback_auto_write)
            self.thread3.start()

        except Exception as e:
            print(e)
            pass
        pass



    def find_gd_clicked(self):
        if self.lineEdit.text() == '':
            self.display.text('请输入工单号！')
            return
        num = int(self.lineEdit.text())
        year = self.comboBox_2.currentText()
        mylist = []
        self.find_btn.setEnabled(False)
        self.thread2 = Search_Gd_Thread(self.Auto, year,num)
        self.thread2._search_str_sign.connect(self.callback_search)
        self.thread2._search_dict_sign.connect(self.callback_search_dict)
        self.thread2.start()


    def glcs_replace(self):
        strText = "完成<font style='background-color:white; color:blue;'>替换</font>字符串!"
        self.display.setText(strText)
        str_find = self.text_find.text().strip()
        str_replace = self.text_replace.text().strip()
        # print(str_find)
        # print(str_replace)
        if (str_find == '' or str_replace == ''):
            return

        row_glcs = self.glcs_table.rowCount()
        if (row_glcs <= 0):
            self.text_find.setText('')
            self.text_replace.setText('')
            return

        for i in range(row_glcs):
            # glcs_list.append(str(self.glcs_table.item(i,2).text()))
            if (self.glcs_table.item(i, 2) == None):   #item  is  NULL
                continue
            temp_str = str(self.glcs_table.item(i,2).text()).strip()    #措施字符串
            temp_str0 = temp_str[:2]        #序号分离，不替换序号
            temp_str1 = temp_str[2:]        #序号后内容
            self.glcs_table.item(i, 2).setBackground(QColor(255, 255, 255))
            if (self.text_find.text().strip() in temp_str1):
                self.glcs_table.setItem(i, 2, QTableWidgetItem(""))
                self.glcs_table.setItem(i, 2, QTableWidgetItem(temp_str0 + temp_str1.replace(str_find, str_replace)))
                self.glcs_table.item(i, 2).setBackground(QColor(255, 255, 0))
        self.text_find.setText('')
        self.text_replace.setText('')
        pass


    #-------隔离措施表格中，插入数据
    def insert_glcs_data(self, glcs_str):
        self.glcs_table.clear()
        if glcs_str.strip() != '无':
            glcs_list = glcs_str.splitlines()
            for i, temp_glcs in enumerate(glcs_list):
                temp_glcs_list = temp_glcs.split('$')
                #删除热机票2 且措施为‘无’
                if ('热机票-2' in temp_glcs_list[0] and '无' in temp_glcs_list[2]):
                    glcs_list.pop(i)
            linesum = len(glcs_list)
            self.glcs_table.setRowCount(linesum)
            self.glcs_table.setColumnCount(3)
            width = self.glcs_table.width()
            self.glcs_table.setColumnWidth(0, width * 0.07)
            self.glcs_table.setColumnWidth(1, width * 0.095)
            self.glcs_table.setColumnWidth(2, width * (1-0.183))
            for index, glcs in enumerate(glcs_list):
                gl_list = glcs.split('$')
                self.glcs_table.setItem(index, 0, QTableWidgetItem(gl_list[0].strip()))
                self.glcs_table.setItem(index, 1, QTableWidgetItem(gl_list[1].strip()))
                self.glcs_table.setItem(index, 2, QTableWidgetItem(gl_list[2].strip()))
        else:
            self.glcs_table.setRowCount(0)
            self.glcs_table.setColumnCount(0)
        pass

    # -------预控措施表格中，插入数据
    def insert_ykcs_data(self, ykcs_str):
        self.ykcs_table.clear()
        if ykcs_str.strip() != '无':
            ykcs_list = ykcs_str.splitlines()
            linesum = len(ykcs_list)
            self.ykcs_table.setRowCount(linesum)
            self.ykcs_table.setColumnCount(2)
            self.ykcs_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
            self.ykcs_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
            for index, ykcs_str in enumerate(ykcs_list):
                yk_list = ykcs_str.split('$')
                newitem0 = QTableWidgetItem(yk_list[0].strip())
                newitem1 = QTableWidgetItem(yk_list[1].strip())
                self.ykcs_table.setItem(index, 0, newitem0)
                self.ykcs_table.setItem(index, 1, newitem1)
        else:
            self.ykcs_table.setRowCount(0)
            self.ykcs_table.setColumnCount(0)
        pass


    #从隔离措施TABLE中，获取数据->拼接成字符串
    def get_glcs_data(self):
        row_glcs = self.glcs_table.rowCount()
        str_glcs = ''
        for i in range(row_glcs):
            strtemp = ''
            stritem0 = self.glcs_table.item(i, 0).text()
            stritem1 = self.glcs_table.item(i, 1).text()
            stritem2 = self.glcs_table.item(i, 2).text()
            strtemp = strtemp + "$" + stritem0 + "$" + stritem1 + "$" + stritem2
            # for j in range(column_glcs):
            #     strtemp = strtemp + "$" + str(self.glcs_table.item(i, j).text())
            str_glcs = str_glcs + strtemp[1:] + '\r\n'
        return  str_glcs

    # 从预控措施TABLE中，获取数据->拼接成字符串
    def get_ykcs_data(self):
        row_ykcs = self.ykcs_table.rowCount()
        column_ykcs = self.ykcs_table.columnCount()
        str_ykcs = ''
        for i in range(row_ykcs):
            strtemp = ''
            for j in range(column_ykcs):
                strtemp = strtemp + "$" + str(self.ykcs_table.item(i, j).text())
            str_ykcs = str_ykcs + strtemp[1:] + '\r\n'
        return str_ykcs


    def clear_info(self):
        self.mvp.setText('')
        self.bianma.setText('')
        self.chengyuan.setText('')
        self.sum.setText('')
        self.glcs_table.clear()
        self.glcs_table.setRowCount(0)
        self.glcs_table.setColumnCount(0)
        self.ykcs_table.clear()
        self.ykcs_table.setRowCount(0)
        self.ykcs_table.setColumnCount(0)



class Login_Thread(QtCore.QThread):
    _login_sign = QtCore.pyqtSignal([int,str])
    def __init__(self,bfs_isintance,employee_info):
        super(Login_Thread,self).__init__()
        self.bfs = bfs_isintance
        self.employee = employee_info
    def __del__(self):
        self.wait()
    def run(self):
        self.bool_login = self.bfs.login(self.employee)
        if self.bool_login:
            self._login_sign.emit(2, '登陆成功')
        else:
            self._login_sign.emit(2, '登陆失败')
class Auto_Write_Thread(QtCore.QThread):
    _auto_write_sign = QtCore.pyqtSignal([int,str])
    def __init__(self,bfs_isintance,aimdict):
        super(Auto_Write_Thread, self).__init__()
        self.bfs = bfs_isintance
        self.aim_dict = aimdict
    def run(self):
        try:
            self._auto_write_sign.emit(2,'打开-->工单页面')
            self.bfs.MyDriver.get(r'http://10.33.2.11/WebBFS/Program3.aspx?programId=216&programNum=10007')
            self._auto_write_sign.emit(2, '切换选项-->工单')
            self.bfs.btn_op('工单')
            self._auto_write_sign.emit(2, '新建-->工单')
            self.bfs.btn_op('new')
            time.sleep(2)
            self._auto_write_sign.emit(2, '写数据-->工单')
            self.bfs.write_page_text('GD', self.aim_dict)

            time.sleep(2)
            self._auto_write_sign.emit(2, '切换选项-->工单分项')
            self.bfs.btn_op('工单分项')
            self._auto_write_sign.emit(2, '写数据-->工单分项')
            self.bfs.write_page_text('GDFX', self.aim_dict)
            time.sleep(2)
            if self.aim_dict['隔离单号'] == '无':
                self._auto_write_sign.emit(2, '定期工作，或无隔离单，操作结束!')
                return
            self._auto_write_sign.emit(2, '切换选项-->连接')
            self.bfs.btn_op('连接')
            self._auto_write_sign.emit(2, '写数据-->连接')
            self.bfs.write_page_text('LJ', self.aim_dict)
            self._auto_write_sign.emit(2, '写数据-->工作票')
            self.bfs.write_page_text('GZP', self.aim_dict)
            time.sleep(2)
            self._auto_write_sign.emit(2, '切换选项-->电厂编码')
            self.bfs.btn_op('电厂编码')
            self._auto_write_sign.emit(2, '写数据-->电厂编码')
            self.bfs.write_page_text('DCBM', self.aim_dict)
            time.sleep(2)
            self._auto_write_sign.emit(2, '切换选项-->隔离措施')
            self.bfs.btn_op('隔离措施')
            self._auto_write_sign.emit(2, '写数据-->隔离措施')
            self.bfs.write_page_text('GLCS', self.aim_dict)
            time.sleep(2)
            self._auto_write_sign.emit(2, '切换选项-->危险预控')
            self.bfs.btn_op('危险预控')
            self._auto_write_sign.emit(2, '写数据-->危险预控')
            self.bfs.write_page_text('YKCS', self.aim_dict)
            self._auto_write_sign.emit(2, '操作完成！')
        except Exception as e:
            self._auto_write_sign.emit(2, e)
        self._auto_write_sign.emit(3, 'END')
        pass

class Search_Gd_Thread(QtCore.QThread):
    _search_str_sign = QtCore.pyqtSignal([int,str])
    _search_dict_sign = QtCore.pyqtSignal([int,dict])
    def __init__(self,bfs_isintance,_year,_num):
        super(Search_Gd_Thread,self).__init__()
        self.bfs = bfs_isintance
        self.year = _year
        self.num = _num
    def run(self):
        try:
            self._search_str_sign.emit(2,'打开-->工单页面')
            self._search_str_sign.emit(1, '5')
            self.bfs.MyDriver.get(r'http://10.33.2.11/WebBFS/Program3.aspx?programId=216&programNum=10007')
            self._search_str_sign.emit(2, '切换选项-->查询')
            self._search_str_sign.emit(1, '10')
            self.bfs.btn_op('查询')
            # time.sleep(3)
            self._search_str_sign.emit(2, '搜索-->工单号')
            self._search_str_sign.emit(1, '15')
            self.bfs.Search_GD(self.num, self.year)
            self._search_str_sign.emit(2, '开始-->查询')
            self._search_str_sign.emit(1, '20')
            self.bfs.btn_op('search')
            mydict = {}
            mydict['工单号'] = self.num
            print(mydict['工单号'])
            print('工单', end='->', flush=True)

            self._search_str_sign.emit(2, '查询-->工单数据')
            self._search_str_sign.emit(1, '25')
            self.bfs.get_page_text('GD', mydict)
            self._search_str_sign.emit(2, '切换选项-->工单分项')
            self._search_str_sign.emit(1, '30')
            self.bfs.btn_op('工单分项')
            print('分项', end='->', flush=True)

            self._search_str_sign.emit(2, '查询-->工单分项数据')
            self._search_str_sign.emit(1, '35')
            self.bfs.get_page_text('GDFX', mydict)
            self._search_str_sign.emit(2, '切换选项-->连接')
            self._search_str_sign.emit(1, '40')
            self.bfs.btn_op('连接')
            print('连接', end='->', flush=True)
            self._search_str_sign.emit(2, '查询-->连接数据')
            self._search_str_sign.emit(1, '45')
            self.bfs.btn_op('选中工作票')
            try:
                element = WebDriverWait(self.bfs.MyDriver, 10).until(
                    EC.visibility_of_element_located((By.ID, 'MC_TC__ctl4_ctl00_LinkViewControl__101')))
            except Exception as e:

                print("获取工作票超时!查询结束....")
                mydict['隔离单号'] = '无'
                mydict['工作票类型'] = '无'
                mydict['总人数'] = '无'
                mydict['班组成员'] = '无'
                mydict['工作内容及地点'] = '无'
                mydict['隔离类型'] = '无'
                mydict['隔离状态'] = '无'
                mydict['隔离类型'] = '无'
                mydict['安全标示'] = '无'
                mydict['隔离措施'] = '无'
                mydict['预控措施'] = '无'
                self._search_str_sign.emit(2, '超时或工作票为空，查询结束！')
                self._search_str_sign.emit(1, '50')
                self._search_dict_sign.emit(3,mydict)
                self._search_str_sign.emit(1, '100')
                self._search_str_sign.emit(3, 'END')
                return

            else:
                qq = self.bfs.findelement("//*[@id='MC_TC__ctl4_ctl00_LinkViewControl__101']")
                num = qq.get_attribute('value')
                mydict['隔离单号'] = num
                print("工作票票号=%d" % int(num), end='->', flush=True)
                self._search_str_sign.emit(2, '工作票存在，开始查询！')
                self._search_str_sign.emit(1, '55')
                self.bfs.btn_op("右键工作票")
                self.bfs.btn_op("打开工作票")

            exist = EC.new_window_is_opened(self.bfs.MyDriver.window_handles)
            WebDriverWait(self.bfs.MyDriver, 20).until(exist)
            self.bfs.MyDriver.switch_to_window(self.bfs.MyDriver.window_handles[-1])
            print('工作票', end='->', flush=True)
            self._search_str_sign.emit(2, '查询-->工作票数据')
            self._search_str_sign.emit(1, '60')
            self.bfs.get_page_text('GZP', mydict)
            self._search_str_sign.emit(2, '切换选项-->电厂编码')
            self._search_str_sign.emit(1, '65')
            self.bfs.btn_op('电厂编码')
            print('电厂编码', end='->', flush=True)
            self._search_str_sign.emit(2, '查询-->电厂编码数据')
            self._search_str_sign.emit(1, '70')
            self.bfs.get_page_text('DCBM', mydict)
            self._search_str_sign.emit(2, '切换选项-->隔离措施')
            self._search_str_sign.emit(1, '75')
            self.bfs.btn_op('隔离措施')
            print('隔离措施', end='->', flush=True)
            self._search_str_sign.emit(2, '查询-->隔离措施数据')
            self._search_str_sign.emit(1, '80')
            self.bfs.get_page_text('GLCS', mydict)
            self._search_str_sign.emit(2, '切换选项-->危险预控')
            self._search_str_sign.emit(1, '85')
            self.bfs.btn_op('危险预控')
            print('危险预控', end='->', flush=True)
            self._search_str_sign.emit(2, '查询-->危险预控数据')
            self._search_str_sign.emit(1, '90')
            self.bfs.get_page_text('WXYK', mydict)
            self.bfs.MyDriver.close()
            self._search_str_sign.emit(2, '查询结束')
            self._search_str_sign.emit(1, '95')
            self.bfs.MyDriver.switch_to_window(self.bfs.MyDriver.window_handles[0])
            self._search_dict_sign.emit(3,mydict)
            self._search_str_sign.emit(1, '100')
        except Exception as e:
            print(e)
            pass
        self._search_str_sign.emit(3, 'END')

        pass





class AutoOperator:
    def __init__(self, isintance):
        # chrome_options = Options()
        # chrome_options.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
        # self.MyDriver=webdriver.Chrome(executable_path=PATH,chrome_options=chrome_options)
        # self.MyDriver=webdriver.Chrome(executable_path=PATH)
        # self.MyDriver.get(r'http://10.33.2.11/WebBFS/Login.aspx')
        # self.input_op(USERNAME,'80180266')
        # self.input_op(PASSWORD,'80180266')
        # self.btn_op('login')
        self.win = isintance
        # self.win.display.setText('登录成功!')
        pass

    def login(self,employee_info_list):
        self.MyDriver = webdriver.Chrome(executable_path=PATH)
        try:
            self.MyDriver.get(r'http://10.33.2.11/WebBFS/Login.aspx')
            self.input_op(USERNAME, employee_info_list[0])
            self.input_op(PASSWORD, employee_info_list[1])
            self.btn_op('login')
        except TimeoutException:
            return False
        return True
        

    def Radio_op(self, geli):
        ele = self.findelement(GZP['电厂编码']['radio_value_电气'])
        value_dianqi = int(ele.get_attribute('value'))
        if value_dianqi == 1:
            print('dianqi===1')
            self.btn_op('radio电气')

        ele = self.findelement(GZP['电厂编码']['radio_value_机械'])
        value_jixie = int(ele.get_attribute('value'))

        if value_jixie == 1:
            print('jixie===1')
            self.btn_op('radio机械')
        if geli == '电气及机械隔离':
            self.btn_op('radio电气')
            self.btn_op('radio机械')
            pass
        elif geli == '电气隔离':
            self.btn_op('radio电气')
        elif geli == '机械隔离':
            self.btn_op('radio机械')

    def Search_GD(self, index, years):  # 查询操作
        self.input_op(GD['查询']['编号'], index)
        self.combo_op(GD['查询']['年份']['combo_btn'], 1)
        self.combo_op(GD['查询']['年份'][years], 2)

    def combo_op(self, xpath, index):  # combo操作
        if index == 1:  # combo控件点击第一步
            ele_exist = EC.presence_of_element_located((By.XPATH, xpath))
            element = WebDriverWait(self.MyDriver, 20).until(ele_exist)
            # element=driver.find_element(By.XPATH,xpath)
            element.click()
        elif index == 2:  # comboxuan选择列表项目
            # ele_exist=EC.presence_of_elements_located((By.XPATH,xpath))[-1]
            # element=WebDriverWait(driver,20).until(ele_exist) 
            element = self.MyDriver.find_elements(By.XPATH, xpath)[
                -1]  # 每次刷新出现3个element,取最后一个,每次点击选取后，增加3个element,刷新后恢复正常3个
            element.click()

    def input_op(self, xpath, content):  # 文本框输入
        ele_exist = EC.presence_of_element_located((By.XPATH, xpath))
        element = WebDriverWait(self.MyDriver, 20).until(ele_exist)
        element.clear()
        element.send_keys(content)

    def read_element_text(self, xpath):  # 读元素数据
        ele_exist = EC.presence_of_element_located((By.XPATH, xpath))
        element = WebDriverWait(self.MyDriver, 20).until(ele_exist)
        return element.get_attribute('value')

    def findelement(self, xpath):
        ele_exist = EC.presence_of_element_located((By.XPATH, xpath))
        return WebDriverWait(self.MyDriver, 10).until(ele_exist)

    def findelements(self, xpath):
        try:
            ele_exist = EC.presence_of_all_elements_located((By.XPATH, xpath))
            return WebDriverWait(self.MyDriver, 10).until(ele_exist)
        except TimeoutException:
            return False

    def get_page_text(self, pagetype, dict):  # 得到玉页面当前数据
        if pagetype == 'GD':
            dict['优先级'] = self.read_element_text(GD['工单']['优先级']['input_adr'])
            dict['维修类型'] = self.read_element_text(GD['工单']['维修类型']['input_adr'])
            dict['负责人'] = self.read_element_text(GD['工单']['负责人']['input_adr'])
            dict['机组'] = self.read_element_text(GD['工单']['机组']['input_adr'])
            dict['专业'] = self.read_element_text(GD['工单']['专业']['input_adr'])
            dict['电厂编码'] = self.read_element_text(GD['工单']['电厂编码']['input_adr'])
            dict['工作内容'] = self.read_element_text(GD['工单']['工作内容']['input_adr'])
            dict['机组类别'] = self.read_element_text(GD['工单']['机组类别']['input_adr'])
            dict['开始日期'] = self.read_element_text(GD['工单']['开始日期']['input_adr'])
            dict['开始时间'] = self.read_element_text(GD['工单']['开始时间']['input_adr'])
            dict['结束日期'] = self.read_element_text(GD['工单']['结束日期']['input_adr'])
            dict['结束时间'] = self.read_element_text(GD['工单']['结束时间']['input_adr'])
            return '工单--完成'
        elif pagetype == 'GDFX':
            qq = self.findelement("//*[@data-ig = 'x:914667624.17:adr:0:tag:']/td[4]")
            dict['班组'] = qq.get_attribute('title')
            return '工单fenxiang--完成'
        elif pagetype == 'GZP':
            dict['工作票类型'] = self.read_element_text(GZP['工作票']['工作票类型']['input_adr'])
            dict['总人数'] = self.read_element_text(GZP['工作票']['总人数']['input_adr'])
            dict['班组成员'] = self.read_element_text(GZP['工作票']['班组成员']['input_adr'])
            dict['工作内容及地点'] = self.read_element_text(GZP['工作票']['工作内容及地点']['input_adr'])
            return '工zuopiao--完成'
        elif pagetype == 'DCBM':
            qq = self.findelement("//*[@data-ig = 'x:1317427084.17:adr:0:tag:']/td[6]")  # GZP['电厂编码']['编码表格']
            dict['隔离类型'] = qq.get_attribute('title')
            return 'bianma--完成'
        elif pagetype == 'GLCS':
            qq1 = self.findelements("//*[@data-ig='x:1317417514.18:mkr:rows:nw:1']/tr")
            if qq1 == False:
                print('无法找到隔离措施!')
                return '隔离措施无法找到！'
            # dict['措施数量']=len(qq1)
            print("隔离措施%d条" % len(qq1), end='->', flush=True)
            str_cs = ''
            for index, q in enumerate(qq1):
                str2 = q.find_element(By.XPATH, "./td[2]")
                str3 = q.find_element(By.XPATH, "./td[3]")
                str5 = q.find_element(By.XPATH, "./td[5]")

                if index == 0:
                    str = ''
                    str9 = q.find_element(By.XPATH, "./td[9]")
                    str10 = q.find_element(By.XPATH, "./td[10]")
                    dict['隔离状态'] = str9.get_attribute('title')
                    dict['安全标示'] = str10.get_attribute('title')
                str = str2.get_attribute('title') + "$" + str5.get_attribute('title') + "$" + str3.get_attribute(
                    'title') + '\r\n'
                str_cs += str
            dict['隔离措施'] = str_cs
            return '隔离措施--完成'

        elif pagetype == 'WXYK':
            qq = self.findelements("//*[@data-ig = 'x:1317421711.16:mkr:rows:nw:1']/tr")
            if qq == False:
                print('无法找到预控措施！')
                return '预控措施无法找到！'

            # dict['预控数量']=len(qq)

            print("预控措施%d条" % len(qq), flush=True)

            str_yk = ''
            for index, q in enumerate(qq):
                str2 = q.find_element(By.XPATH, "./td[2]")
                str3 = q.find_element(By.XPATH, "./td[3]")
                str = str2.get_attribute('title') + "$" + str3.get_attribute('title') + '\r\n'
                str_yk = str_yk + str
            dict['预控措施'] = str_yk
            return '预控措施--完成'

    def get_gd_num(self, xpath):
        qq = self.findelements(xpath)
        num = len(qq)
        gd_num = []
        for q in qq:
            str1 = q.find_element(By.XPATH, "./td[1]")
            gd_num.append(str1.get_attribute('title'))
        return gd_num

    def write_page_text(self, pagetype, aim_dict):  # 写入页面数据
        if pagetype == 'GD':
            str1 = aim_dict['优先级']
            str2 = aim_dict['维修类型']
            str3 = aim_dict['机组']
            str4 = aim_dict['专业']
            str5 = aim_dict['机组类别']
            self.combo_op(GD['工单']['优先级']['combo_btn'], 1)
            self.combo_op(GD['工单']['优先级'][str1], 2)
            print('优先级..........以写入')
            # time.sleep(1)    
            self.combo_op(GD['工单']['维修类型']['combo_btn'], 1)
            self.combo_op(GD['工单']['维修类型'][str2], 2)
            print('维修类型..........以写入')
            # time.sleep(1)
            self.combo_op(GD['工单']['机组']['combo_btn'], 1)
            self.combo_op(GD['工单']['机组'][str3], 2)

            print('机组..........以写入')
            # time.sleep(1)
            self.combo_op(GD['工单']['专业']['combo_btn'], 1)
            self.combo_op(GD['工单']['专业'][str4], 2)
            print('专业..........以写入')
            # time.sleep(1)
            self.combo_op(GD['工单']['机组类别']['combo_btn'], 1)
            self.combo_op(GD['工单']['机组类别'][str5], 2)
            print('机组类别..........以写入')
            self.input_op(GD['工单']['负责人']['input_adr'], aim_dict['负责人'])
            print('负责人..........以写入')  # 负责人
            self.input_op(GD['工单']['工作内容']['input_adr'], aim_dict['工作内容'])  # 工作内容
            print('工作内容..........以写入')
            self.input_op(GD['工单']['开始日期']['input_adr'], aim_dict['开始日期'])  # 开始日期
            print('开始日期..........以写入')
            self.input_op(GD['工单']['开始时间']['input_adr'], aim_dict['开始时间'])  # 开始时间
            print('开始时间..........以写入')
            self.input_op(GD['工单']['结束日期']['input_adr'], aim_dict['结束日期'])  # 结束日期
            print('结束日期..........以写入')
            self.input_op(GD['工单']['结束时间']['input_adr'], aim_dict['结束时间'])  # 结束时间
            print('结束时间..........以写入')
            self.input_op(GD['工单']['电厂编码']['input_adr'], aim_dict['电厂编码'])  # 电厂编码
            print('电厂编码..........以写入')
            ele = self.MyDriver.find_element(By.ID, 'MC_TC__ctl2_ctl00__301')  # 点空白处 加载电厂编码
            ele.click()
            time.sleep(3)
            print('*' * 60)
            self.btn_op('save')
        elif pagetype == 'GDFX':
            self.btn_op('分项表格')
            time.sleep(5)
            self.input_op(GD['工单分项']['班组'], aim_dict['班组'])  # aim_dict['班组']
            self.btn_op('save')
        elif pagetype == 'LJ':
            self.btn_op('选中工作票')
            time.sleep(1)
            self.btn_op('右键工作票')
            time.sleep(1)
            self.btn_op('创建工作票')
            time.sleep(3)
            self.btn_op('选中工作票')
            time.sleep(1)
            self.btn_op('右键工作票')
            time.sleep(1)
            self.btn_op('打开工作票')
            exist = EC.new_window_is_opened(self.MyDriver.window_handles)
            WebDriverWait(self.MyDriver, 20).until(exist)
            # self.MyDriver.switch_to_window(self.MyDriver.window_handles[-1])
            self.MyDriver.switch_to.window(self.MyDriver.window_handles[-1])

            pass
        elif pagetype == 'GZP':
            str1 = aim_dict['工作票类型']
            str2 = aim_dict['专业']
            str3 = aim_dict['机组']
            self.combo_op(GZP['工作票']['工作票类型']['combo_btn'], 1)
            self.combo_op(GZP['工作票']['工作票类型'][str1], 2)
            print('工作票类型..........以写入')
            self.combo_op(GZP['工作票']['专业']['combo_btn'], 1)
            self.combo_op(GZP['工作票']['专业'][str2], 2)
            print('专业..........以写入')
            self.combo_op(GZP['工作票']['机组']['combo_btn'], 1)
            self.combo_op(GZP['工作票']['机组'][str3], 2)
            print('机组..........以写入')
            self.input_op(GZP['工作票']['负责人']['input_adr'], aim_dict['负责人'])
            print('负责人..........以写入')
            self.input_op(GZP['工作票']['总人数']['input_adr'], aim_dict['总人数'])
            print('负责人..........以写入')
            self.input_op(GZP['工作票']['班组成员']['input_adr'], aim_dict['班组成员'])
            print('班组成员..........以写入')
            self.input_op(GZP['工作票']['工作内容及地点']['input_adr'], aim_dict['工作内容及地点'])
            print('工作内容及地点..........以写入')
            print('*' * 60)
            self.btn_op('save')
        elif pagetype == 'DCBM':
            self.btn_op('编码表格')
            str3 = aim_dict.get('隔离类型')
            if str3 == '' or str3 == 'nan':
                return
            print(str3)
            time.sleep(5)
            self.combo_op(GZP['电厂编码']['隔离']['combo_btn'], 1)
            self.combo_op(GZP['电厂编码']['隔离'][str3], 2)
            self.Radio_op(str3)
            self.btn_op('save')

        elif pagetype == 'GLCS':
            tempstr = str(aim_dict['隔离措施'])
            glcs_list = tempstr.splitlines()
            print(glcs_list)
            print(aim_dict['安全标识'])
            for index, glcs in enumerate(glcs_list):
                str1 = glcs.split('$')
                if str1[0].strip() == '热机票-2' or str1[0].strip() == '' :
                    break
                if index == 0:
                    self.combo_op(GZP['隔离措施']['隔离切换操作']['combo_btn'], 1)
                    self.combo_op(GZP['隔离措施']['隔离切换操作']['隔离切换操作'], 2)
                    self.combo_op(GZP['隔离措施']['措施类别']['combo_btn'], 1)
                    self.combo_op(GZP['隔离措施']['措施类别']['热机票-1'], 2)
                    self.input_op(GZP['隔离措施']['措施内容'], str1[2])
                    self.input_op(GZP['隔离措施']['电厂编码'], aim_dict['电厂编码'])
                    qq = self.MyDriver.find_element(By.XPATH, GZP['隔离措施']['措施内容'])
                    qq.click()
                    time.sleep(2)
                    if aim_dict['隔离状态'] != '无':
                        self.combo_op(GZP['隔离措施']['隔离状态']['combo_btn'], 1)
                        self.combo_op(GZP['隔离措施']['隔离状态'][aim_dict['隔离状态']], 2)
                    if aim_dict['安全标识'] != '无':
                        self.combo_op(GZP['隔离措施']['安全标示']['combo_btn'], 1)
                        self.combo_op(GZP['隔离措施']['安全标示'][aim_dict['安全标识']], 2)
                    self.btn_op('save')
                    time.sleep(3)
                else:
                    self.combo_op(GZP['隔离措施']['隔离切换操作']['combo_btn'], 1)
                    self.combo_op(GZP['隔离措施']['隔离切换操作']['记录陈述'], 2)
                    time.sleep(1)
                    self.combo_op(GZP['隔离措施']['措施类别']['combo_btn'], 1)
                    self.combo_op(GZP['隔离措施']['措施类别']['热机票-1'], 2)
                    self.input_op(GZP['隔离措施']['措施内容'], str1[2])
                    self.btn_op('save')
                    time.sleep(2)
            pass
        elif pagetype == 'YKCS':
            tempstr = str(aim_dict['危险预控'])
            ykcs_list = tempstr.splitlines()
            print(ykcs_list)
            for index, glcs in enumerate(ykcs_list):
                str1 = glcs.split('$')
                self.input_op(GZP['危险预控']['危险辨识'], str1[0])
                self.input_op(GZP['危险预控']['预控措施'], str1[1])
                self.btn_op('save')
                time.sleep(2)


    def btn_op(self, operator):  # 按钮操作
        def btn_click(xpath):
            ele_exist = EC.presence_of_element_located((By.XPATH, xpath))
            element = WebDriverWait(self.MyDriver, 20).until(ele_exist)
            element.click()

        if operator == '工单':            btn_click(GD['工单']['TAB_btn'])
        elif operator == '工单分项':       btn_click(GD['工单分项']['TAB_btn'])
        elif operator == '连接':          btn_click(GD['连接']['TAB_btn'])
        elif operator == '工作票':         btn_click(GZP['工作票']['TAB_btn'])
        elif operator == '电厂编码':        btn_click(GZP['电厂编码']['TAB_btn'])
        elif operator == '隔离措施':        btn_click(GZP['隔离措施']['TAB_btn'])
        elif operator == '危险预控':        btn_click(GZP['危险预控']['TAB_btn'])
        elif operator == 'search':         btn_click(BTN_Search)
        elif operator == 'save':           btn_click(BTN_Save)
        elif operator == 'new':            btn_click(BTN_New)
        elif operator == 'submit1':        btn_click(BTN_Submit1)
        elif operator == 'submit2':        btn_click(BTN_Submit2)
        elif operator == '分项表格':        btn_click(GD['工单分项']['分项表格'])
        elif operator == '编码表格':        btn_click(GZP['电厂编码']['编码表格'])
        elif operator == '分项确认':        btn_click(BTN_ACK)
        elif operator == '选中工作票':      btn_click(GD['连接']['工作票'])
        elif operator == '打开工作票':      btn_click(GD['连接']['打开工作票'])
        elif operator == '创建工作票':      btn_click(GD['连接']['创建工作票'])
        elif operator == 'radio电气':      btn_click(GZP['电厂编码']['radio_btn_电气'])
        elif operator == 'radio机械':      btn_click(GZP['电厂编码']['radio_btn_机械'])
        elif operator == '右键工作票':
            ele_exist = EC.presence_of_element_located((By.XPATH, GD['连接']['工作票']))
            element = WebDriverWait(self.MyDriver, 20).until(ele_exist)
            (ActionChains)(self.MyDriver).context_click(element).perform()
        elif operator == 'login':          btn_click(LOGIN)


if __name__ == "__main__":
    app = QApplication([])
    mywindow = Window()
    mywindow.show()
    app.exec_()
