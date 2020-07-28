from PyQt5.QtWidgets import QWidget, QApplication, QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView, QComboBox, QLabel
from PyQt5.QtCore import QDate, QTime, QStringListModel
from PyQt5.QtGui import QIntValidator
from  PyQt5.QtGui import QColor
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from ui_easy1 import *
from mydict import GD, GZP, EMPLOYEE_INFO

BTN_Save = "//*[@id='ctl00xcphToolbarxuwtToolbar_Item_4_img']"
BTN_Search = "//*[@id='ctl00xcphToolbarxuwtToolbar_Item_5_img']"
BTN_New = "//*[@id='ctl00xcphToolbarxuwtToolbar_Item_19_img']"
BTN_Submit1 = "//*[@id='MC_TC__ctl2_ctl00__964']"
BTN_Submit2 = "//*[@id='MC_TC__ctl2_ctl00__5025']"
BTN_ACK = "//*[@id='MC_TC__ctl3_ctl00_Splitter_10007_4_1_2__1985']"
USERNAME = '//*[@name="tbUser"]'
PASSWORD = '//*[@name="tbPassword"]'
LOGIN = '//*[@name="btnLogin"]'
FILEPATH = r'data\easy.xlsx'
PATH = r"d:\app\chromedriver.exe"
AIMDICT = {}
# OPERATE_PAGE = ''


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
        self.btn_glcs_addrow.clicked.connect(self.add_glcs_rows)
        self.btn_glcs_delrow.clicked.connect(self.remove_glcs_rows)
        self.btn_ykcs_addrow.clicked.connect(self.add_ykcs_rows)
        self.btn_ykcs_delrow.clicked.connect(self.remove_ykcs_rows)
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

        self.ceshi.clicked.connect(self.test111)
    def test111(self):
        self.Auto = AutoOperator(self)
        print('test')
        self.Auto.test_login()
        print(self.Auto)
        AIMDICT = {}
        self.get_widget_info(AIMDICT)

        self.Auto.write_page_text('GLCS', AIMDICT)
        print('123')
        pass

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
            self.insert_data(df)

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

        # self.Auto.login('123')
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
                self.current_dataframe = self.first_dataframe
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
        if self.mvp.text() == '' or self.bianma.text() == '' or self.chengyuan.text() == ''\
                or self.sum.text() == '' or self.content.text() == '' or self.contentaddr.text() == '':
            print('有内容为空！')
        else:
            aim_dict = {}
            self.get_widget_info(aim_dict)
            aim_dict['序号'] = self.xuhao_last + 1
            mylist = []
            mylist.append(aim_dict)
            df = pd.DataFrame(mylist)

            try:
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
            except Exception as e:
                print('保存错误:' + str(e))
            else:
                print('保存成功！')
                self.get_df_from_excel(FILEPATH)
            finally:
                self.is_edit = False




    def del_dataframe(self):  #删除选中data
        if self.mvp.text() == '':
            print('没有选中ITEM')
            return
        else:

            del_list = []
            # df = self.current_dataframe[self.current_dataframe['序号'].isin([self.select_xuhao])]
            # 遍历df找到序号对应的索引
            try:
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
            except Exception as e:
                print("删除错误：" + str(e))


    #通过年份，获取主dateframe数据，并插入数据到主TABLE
    def get_df_from_excel(self, file_path, years='模板'):
        self.first_dataframe = pd.read_excel(file_path)
        self.current_dataframe = self.first_dataframe
        #获取最后一行序号
        self.insert_data(self.current_dataframe)
        # self.xuhao_last = self.first_dataframe.tail(1)['序号'].values[0]
        # print('最后序号为{}'.format(self.xuhao_last))



    #隔离措施TABLE增加行
    def add_glcs_rows(self):
        self.is_edit = True
        currentrow = self.glcs_table.currentRow()
        rowcount = self.glcs_table.rowCount()
        if currentrow == -1:
            self.glcs_table.insertRow(rowcount)
        else:
            self.glcs_table.insertRow(currentrow + 1)


    # 隔离措施TABLE删除行
    def remove_glcs_rows(self):
        self.is_edit = True
        selections = self.glcs_table.selectionModel()
        selectionslist = selections.selectedRows()
        rows = []
        for r in selectionslist:
            rows.append(r.row())
        rows.reverse()
        for i in rows:
            self.glcs_table.removeRow(i)


    # 预控措施TABLE增加行
    def add_ykcs_rows(self):
        self.is_edit = True
        currentrow = self.ykcs_table.currentRow()
        rowcount = self.ykcs_table.rowCount()
        if currentrow == -1:
            self.ykcs_table.insertRow(rowcount)
        else:
            self.ykcs_table.insertRow(currentrow + 1)

        # 预控措施TABLE删除行

    def remove_ykcs_rows(self):
        self.is_edit = True
        selections = self.ykcs_table.selectionModel()
        selectionslist = selections.selectedRows()
        rows = []
        for r in selectionslist:
            rows.append(r.row())
        rows.reverse()
        for i in rows:
            self.ykcs_table.removeRow(i)


    #插入数据到主TABLE
    def insert_data(self, current_df):
        try:
            self.xuhao_last = current_df.tail(1)['序号'].values[0]
            # print('最后序号为{}'.format(self.xuhao_last))
            dingqi, weixiu = self.filter_dateframe(current_df, type='dingqi-fenli')
            if self.checkBox.isChecked():
                temp_df = dingqi
            else:
                temp_df = weixiu
            selected_df = temp_df[['序号',  '负责人', '机组', '工作内容', '维修类型',]]
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
            # self.main_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
            self.main_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
            self.main_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
            for i in range(rows):
                for j in range(columns):
                    newitem = QTableWidgetItem(str(selected_df.iat[i, j]))
                    self.main_table.setItem(i, j, newitem)
        except Exception as e:
            print(e)



    #主TABLE点击事件
    def main_table_clicked(self, item):
        if self.is_edit == True:
            if self.select_xuhao != None:
                res = QMessageBox.information(None, 'title', '信息已改变，是否离开', QMessageBox.Yes | QMessageBox.No)
                if res == QMessageBox.No:
                    self.main_table.selectRow(self.selected_index)
                    return
            self.is_edit = False
        self.content.setStyleSheet("background-color: rgb(252, 228, 189);")
        self.contentaddr.setStyleSheet("background-color: rgb(252, 228, 189);")
        self.jizu.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.jizuleibie.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.bianma.setStyleSheet("background-color: rgb(255, 255, 255);")
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
        # print(f(df['维修类型']))
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


    def get_widget_info(self, out_dict):
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
        if self.Auto == None:
            QMessageBox.information(None, "title", "没有对象！",QMessageBox.Yes)
            return
        # else:
        #     QMessageBox.information(None, "title", "YES！", QMessageBox.Yes)
        #     return
        try:
            str_glcs = self.get_glcs_data()
            str_ykcs = self.get_ykcs_data()
            AIMDICT = {}
            self.get_widget_info(AIMDICT)
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
            # AIMDICT['安全标示'] = self.anquanbiaoshi.currentText()
            # AIMDICT['隔离措施'] = str_glcs
            # AIMDICT['预控措施'] = str_ykcs
            # print(AIMDICT)
            # return

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
        content = self.content.text()
        contentaddr = self.contentaddr.text()
        if str_find in content:
            content = content.replace(str_find, str_replace)
            self.content.setText(content)
            self.content.setStyleSheet("background-color: rgb(255, 255, 0);")
            self.is_edit = True

        if str_find in contentaddr:
            contentaddr = contentaddr.replace(str_find, str_replace)
            self.contentaddr.setText(contentaddr)
            self.contentaddr.setStyleSheet("background-color: rgb(255, 255, 0);")
            self.is_edit = True

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
        self.jizu.setStyleSheet("background-color: rgb(255, 0, 0);")
        self.jizuleibie.setStyleSheet("background-color: rgb(255, 0, 0);")
        self.bianma.setStyleSheet("background-color: rgb(255, 0, 0);")
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
        #---------------------------------test----------------------------------
        def btn_click(xpath):
            ele_exist = EC.presence_of_element_located((By.XPATH, xpath))
            element = WebDriverWait(self.bfs.MyDriver, 20).until(ele_exist)
            element.click()

        def input_op(xpath, content):  # 文本框输入
            ele_exist = EC.presence_of_element_located((By.XPATH, xpath))
            element = WebDriverWait(self.bfs.MyDriver, 20).until(ele_exist)
            element.clear()
            element.send_keys(content)

        保存 = "//*[@id='ctl00xcphToolbarxuwtToolbar_Item_4_img']"
        班组 = r"//*[@id='MC_TC__ctl3_ctl00_Splitter_10007_4_1_3__960']"
        分项表格 = r"//*[@data-ig = 'x:914667624.17:adr:0:tag:']"
        # 分项 = "//td[@tabid='MC_TC,3']"
        # btn_click(分项)
        # btn_click(分项表格)
        # time.sleep(5)
        # input_op(班组, "热控计算机")
        # btn_click(保存)
        # time.sleep(2)
        # btn_click(GD['连接']['TAB_btn'])






        #-----------------------------------test------------------------------------

        #7-18 注视，需要恢复
        try:
            self._auto_write_sign.emit(2,'打开-->工单页面')
            self.bfs.MyDriver.get(r'http://10.33.2.11/WebBFS/Program3.aspx?programId=216&programNum=10007')

            self._auto_write_sign.emit(2, '切换选项-->工单')
            self.bfs.btn_op('工单')

            self._auto_write_sign.emit(2, '新建-->工单')
            time.sleep(5)
            self.bfs.btn_op('new')
            time.sleep(5)

            self._auto_write_sign.emit(2, '写数据-->工单')
            self.bfs.write_page_text('GD', self.aim_dict)
            time.sleep(5)

            self._auto_write_sign.emit(2, '切换选项-->工单分项')
            self.bfs.btn_op('工单分项')

            self._auto_write_sign.emit(2, '写数据-->工单分项')
            # self.bfs.write_page_text('GDFX', self.aim_dict)
            btn_click(分项表格)
            time.sleep(5)
            input_op(班组, "热控计算机")
            tic = time.time()
            while 1:
                btn_click(保存)
                cur_tic = time.time()
                if cur_tic - tic >8:
                    break
                time.sleep(1)
            time.sleep(4)

            if self.aim_dict['隔离单号'] == '无':
                self._auto_write_sign.emit(2, '定期工作，或无隔离单，操作结束!')
                return

            self._auto_write_sign.emit(2, '切换选项-->连接')
            self.bfs.btn_op('连接')

            self._auto_write_sign.emit(2, '写数据-->连接')
            self.bfs.write_page_text('LJ', self.aim_dict)

            self._auto_write_sign.emit(2, '写数据-->工作票')
            self.bfs.write_page_text('GZP', self.aim_dict)
            time.sleep(5)

            self._auto_write_sign.emit(2, '切换选项-->电厂编码')
            self.bfs.btn_op('电厂编码')

            self._auto_write_sign.emit(2, '写数据-->电厂编码')
            self.bfs.write_page_text('DCBM', self.aim_dict)
            time.sleep(5)

            self._auto_write_sign.emit(2, '切换选项-->隔离措施')
            self.bfs.btn_op('隔离措施')

            self._auto_write_sign.emit(2, '写数据-->隔离措施')
            self.bfs.write_page_text('GLCS', self.aim_dict)
            time.sleep(5)

            self._auto_write_sign.emit(2, '切换选项-->预控措施')
            self.bfs.btn_op('预控措施')

            self._auto_write_sign.emit(2, '写数据-->预控措施')
            self.bfs.write_page_text('YKCS', self.aim_dict)

            self._auto_write_sign.emit(2, '操作完成！')
        except Exception as e:
            self._auto_write_sign.emit(2, str(e))
        self._auto_write_sign.emit(3, 'END')
        # pass


        #  7-18 注视，需要恢复


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

            # exist = EC.new_window_is_opened(self.bfs.MyDriver.window_handles)
            # WebDriverWait(self.bfs.MyDriver, 20).until(exist)
            time.sleep(3)
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
            self._search_str_sign.emit(2, '切换选项-->预控措施')
            self._search_str_sign.emit(1, '85')
            self.bfs.btn_op('预控措施')
            print('预控措施', end='->', flush=True)
            self._search_str_sign.emit(2, '查询-->预控措施数据')
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
        self.win = isintance
        self.OPERATE_PAGE = ''

    def test_login(self):
        PATH1 = r'c:\chromedriver.exe'
        chrome_options = Options()
        chrome_options.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
        self.MyDriver = webdriver.Chrome(executable_path=PATH1, chrome_options=chrome_options)
        return True

    def login(self,employee_info_list):
        #----------------huifu-------------------
        binary = FirefoxBinary(r'C:\WebBFS\App\Firefox\firefox.exe')
        self.MyDriver = webdriver.Firefox(firefox_binary = binary)
        try:
            self.MyDriver.get(r'http://10.33.2.11/WebBFS/Login.aspx')
            self.input_op(USERNAME, employee_info_list[0])
            self.input_op(PASSWORD, employee_info_list[1])
            self.btn_op('login')
        except TimeoutException:
            return False
        return True
        #----------------huifu---------------

        #------------------test-----------------
        # PATH1 = r'c:\chromedriver.exe'
        # chrome_options = Options()
        # chrome_options.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
        # self.MyDriver = webdriver.Chrome(executable_path=PATH1, chrome_options=chrome_options)
        #
        # return True
        # ------------------test-----------------

    def Radio_op(self, geli):
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
            element = self.MyDriver.find_elements(By.XPATH, xpath)[-1]  # 每次刷新出现3个element,取最后一个,每次点击选取后，增加3个element,刷新后恢复正常3个
            element.click()


    def combobox_operation(self, combo, operation):
        if self.OPERATE_PAGE == "GD工单":
            combo_xpath = GD['工单'][combo]['combo_btn']
            combolist_xpath = GD['工单'][combo][operation]
        elif self.OPERATE_PAGE == "GD工单分项":
            combo_xpath = GD['工单分项'][combo]['combo_btn']
            combolist_xpath = GD['工单分项'][combo][operation]
        elif self.OPERATE_PAGE == "GD连接":
            pass
        elif self.OPERATE_PAGE == "GZP工作票":
            combo_xpath = GZP['工作票'][combo]['combo_btn']
            combolist_xpath = GZP['工作票'][combo][operation]
        elif self.OPERATE_PAGE == "GZP电厂编码":
            combo_xpath = GZP['电厂编码'][combo]['combo_btn']
            combolist_xpath = GZP['电厂编码'][combo][operation]
        elif self.OPERATE_PAGE == "GZP隔离措施":
            combo_xpath = GZP['隔离措施'][combo]['combo_btn']
            combolist_xpath = GZP['隔离措施'][combo][operation]
        elif self.OPERATE_PAGE == "GZP预控措施":
            combo_xpath = GZP['预控措施'][combo]['combo_btn']
            combolist_xpath = GZP['预控措施'][combo][operation]
        else:
            print('self.OPERATE_PAGE没有找到！')
            return
        count = 0
        while count < 5:
            try:
                if count > 0:
                    print(operation + "---重复执行X{}".format(count))
                ele_exist = EC.presence_of_element_located((By.XPATH, combo_xpath))
                element = WebDriverWait(self.MyDriver, 5).until(ele_exist)
                element.click()
                element = self.MyDriver.find_elements(By.XPATH, combolist_xpath)[-1]
                element.click()
            except Exception as e:
                print(operation + "---error")
                print(e)
                count += 1
                time.sleep(1)
            else:
                break

    def input_operation(self, input, content):
        if self.OPERATE_PAGE == "GD工单":
            input_xpath = GD['工单'][input]['input_adr']
        elif self.OPERATE_PAGE == "GD工单分项":
            input_xpath = GD['工单分项'][input]['input_adr']
        elif self.OPERATE_PAGE == "GD连接":
            pass
        elif self.OPERATE_PAGE == "GZP工作票":
            input_xpath = GZP['工作票'][input]['input_adr']
        elif self.OPERATE_PAGE == "GZP电厂编码":
            input_xpath = GZP['电厂编码'][input]['input_adr']
        elif self.OPERATE_PAGE == "GZP隔离措施":
            input_xpath = GZP['隔离措施'][input]['input_adr']
        elif self.OPERATE_PAGE == "GZP预控措施":
            input_xpath = GZP['预控措施'][input]['input_adr']
        else:
            print('self.OPERATE_PAGE没有找到！')
            return

        count = 0
        while count < 5:
            try:
                if count > 0:
                    print(input + "---重复执行X{}".format(count))
                ele_exist = EC.presence_of_element_located((By.XPATH, input_xpath))
                element = WebDriverWait(self.MyDriver, 5).until(ele_exist)
                element.clear()
                element.send_keys(content)
            except Exception as e:
                print(input + "---error")
                print(e)
                count += 1
                time.sleep(1)
            else:
                break



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

    # def get_gd_num(self, xpath):
    #     qq = self.findelements(xpath)
    #     num = len(qq)
    #     gd_num = []
    #     for q in qq:
    #         str1 = q.find_element(By.XPATH, "./td[1]")
    #         gd_num.append(str1.get_attribute('title'))
    #     return gd_num

    def write_page_text(self, pagetype, aim_dict):  # 写入页面数据
        def save_plus(tic = time.time()):
            count = 1
            while 1:
                cur_tic = time.time()
                if cur_tic - tic > 5:
                    break
                self.btn_op('save')
                print("第{}次保存！".format(count))
                count += 1
                time.sleep(1)

        if pagetype == 'GD':
            self.OPERATE_PAGE = 'GD工单'

            print(self.OPERATE_PAGE)
            self.combobox_operation('优先级', aim_dict['优先级']);      print('优先级..........已写入')
            self.combobox_operation('维修类型', aim_dict['维修类型']);  print('维修类型..........已写入')
            self.combobox_operation('机组', aim_dict['机组']);           print('机组..........已写入')
            self.combobox_operation('专业', aim_dict['专业']);           print('专业..........已写入')
            self.combobox_operation('机组类别', aim_dict['机组类别']);   print('机组类别..........已写入')
            self.input_operation('负责人', aim_dict['负责人']);          print('负责人..........已写入')
            self.input_operation('工作内容', aim_dict['工作内容']);      print('工作内容..........已写入')
            self.input_operation('开始日期', aim_dict['开始日期']);      print('开始日期..........已写入')
            self.input_operation('开始时间', aim_dict['开始时间']);      print('开始时间..........已写入')
            self.input_operation('结束日期', aim_dict['结束日期']);      print('结束日期..........已写入')
            self.input_operation('结束时间', aim_dict['结束时间']);      print('结束时间..........已写入')
            self.input_operation('电厂编码', aim_dict['电厂编码']);      print('电厂编码..........已写入')
            ele = self.MyDriver.find_element(By.ID, 'MC_TC__ctl2_ctl00__301')  # 点空白处 加载电厂编码
            ele.click()
            while 1:
                x = self.MyDriver.find_element_by_id('MC_TC__ctl2_ctl00__337')   #电厂编码 名称
                x_name = str(x.get_attribute('value'))
                if x_name.strip() != "":
                    time.sleep(1)
                    break
            time.sleep(1)
            print('*' * 60)
            tic = time.time()
            save_plus(tic)

        elif pagetype == 'GDFX':
            self.OPERATE_PAGE = 'GD工单分项'
            self.btn_op('分项表格')
            time.sleep(5)
            self.input_operation('班组', aim_dict['班组']);         print('班组..........已写入')

            tic = time.time()
            save_plus(tic)

        elif pagetype == 'LJ':
            self.OPERATE_PAGE = 'GD连接'
            self.btn_op('选中工作票')
            time.sleep(1)
            self.btn_op('右键工作票')
            time.sleep(1)
            self.btn_op('创建工作票')
            print('创建工作票')
            time.sleep(3)
            self.btn_op('选中工作票')
            time.sleep(1)
            self.btn_op('右键工作票')
            time.sleep(1)
            self.btn_op('打开工作票')
            print('打开工作票')
            # exist = EC.new_window_is_opened(self.MyDriver.window_handles)
            # WebDriverWait(self.MyDriver, 20).until(exist)
            time.sleep(5)
            self.MyDriver.switch_to.window(self.MyDriver.window_handles[-1])

            pass
        elif pagetype == 'GZP':
            self.OPERATE_PAGE = 'GZP工作票'

            self.combobox_operation('工作票类型', aim_dict['工作票类型']);   print('工作票类型..........已写入')
            self.combobox_operation('专业', aim_dict['专业']);                print('专业..........已写入')
            self.combobox_operation('机组', aim_dict['机组']);                print('机组..........已写入')
            # self.input_op(GZP['工作票']['负责人']['input_adr'], aim_dict['负责人'])
            self.input_operation('负责人', aim_dict['负责人']);               print('负责人..........已写入')
            self.input_operation('总人数', aim_dict['总人数']);               print('总人数..........已写入')
            self.input_operation('班组成员', aim_dict['班组成员']);           print('班组成员..........已写入')
            self.input_operation('工作内容及地点', aim_dict['工作内容及地点']); print('工作内容及地点..........已写入')
            print('*' * 60)
            save_plus(tic = time.time())

        elif pagetype == 'DCBM':
            self.OPERATE_PAGE = 'GZP电厂编码'
            self.btn_op('编码表格')
            if aim_dict.get('隔离类型') == '' or aim_dict.get('隔离类型') == 'nan':
                return
            # print(str3)
            time.sleep(5)
            self.combobox_operation('隔离', aim_dict.get('隔离类型'))
            self.Radio_op(aim_dict.get('隔离类型'))
            save_plus(tic = time.time())

        elif pagetype == 'GLCS':
            self.OPERATE_PAGE = 'GZP隔离措施'
            tempstr = str(aim_dict['隔离措施'])
            glcs_list = tempstr.splitlines()
            len1 = len(glcs_list)
            for index, glcs in enumerate(glcs_list):
                str1 = glcs.split('$')
                print(str1)
                if str1[0].strip() == '热机票-2' or str1[0].strip() == '' :
                    break
                print("隔离措施一共{}条，正在写入第{}条！".format(len1, index + 1))
                if index == 0:

                    self.combobox_operation('隔离切换操作', '隔离切换操作')
                    self.combobox_operation('措施类别', '热机票-1')
                    print(GZP['隔离措施']['措施内容']['input_adr'])
                    self.input_operation('措施内容', str1[2])
                    print(str1[2])
                    self.input_operation('电厂编码', aim_dict['电厂编码'])
                    qq = self.MyDriver.find_element(By.XPATH, GZP['隔离措施']['措施内容']['input_adr'])
                    qq.click()
                    time.sleep(2)
                    if aim_dict['隔离状态'] != '无':
                        self.combobox_operation('隔离状态', aim_dict['隔离状态'])
                    if aim_dict['安全标示'] != '无':
                        self.combobox_operation('安全标示', aim_dict['安全标示'])
                    self.btn_op('save')
                    time.sleep(3)
                else:
                    self.combobox_operation('隔离切换操作', '记录陈述')


                    time.sleep(1)
                    self.combobox_operation('措施类别', '热机票-1')
                    self.input_operation('措施内容', str1[2])
                    self.btn_op('save')
                    time.sleep(2)
            pass
        elif pagetype == 'YKCS':
            self.OPERATE_PAGE = 'GZP预控措施'
            tempstr = str(aim_dict['预控措施'])
            ykcs_list = tempstr.splitlines()
            len2 = len(ykcs_list)
            for index, glcs in enumerate(ykcs_list):
                print("预控措施一共{}条，正在写入第{}条！".format(len2, index + 1))
                str1 = glcs.split('$')
                self.input_operation('危险辨识',str1[0])
                self.input_operation('预控措施', str1[1])
                self.btn_op('save')
                time.sleep(2)


    def btn_op(self, operator):  # 按钮操作
        def btn_click(xpath):
            ele_exist = EC.presence_of_element_located((By.XPATH, xpath))
            element = WebDriverWait(self.MyDriver, 5).until(ele_exist)
            element.click()

        def toolbar_click( xpath):
            ele_exist = EC.element_to_be_clickable((By.XPATH, xpath))
            element = WebDriverWait(self.MyDriver, 5).until(ele_exist)
            element.click()

        if operator == '工单':            btn_click(GD['工单']['TAB_btn'])
        elif operator == '工单分项':       btn_click(GD['工单分项']['TAB_btn'])
        elif operator == '连接':          btn_click(GD['连接']['TAB_btn'])
        elif operator == '工作票':         btn_click(GZP['工作票']['TAB_btn'])
        elif operator == '电厂编码':        btn_click(GZP['电厂编码']['TAB_btn'])
        elif operator == '隔离措施':        btn_click(GZP['隔离措施']['TAB_btn'])
        elif operator == '预控措施':        btn_click(GZP['预控措施']['TAB_btn'])
        elif operator == 'search':         toolbar_click(BTN_Search)
        elif operator == 'save':            toolbar_click(BTN_Save)
        elif operator == 'new':            toolbar_click(BTN_New)
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
