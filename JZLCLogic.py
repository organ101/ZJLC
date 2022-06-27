from PyQt5.QtWidgets import QWizard, QFileDialog,QMessageBox
from JZLCwizardGUI import Ui_Wizard
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import os
import pandas as pd
import traceback

console_plat = sys.stdout  #保存之前的sys.stdout，便于之后恢复到控制台显示print信息

MEMBER_MAP_DICT ={
    "湖州银行股份有限公司": "是",
    "嘉兴银行股份有限公司": "是",
    "绍兴银行股份有限公司": "是",
    "台州银行股份有限公司": "是",
    "金华银行股份有限公司": "是",
    "温州银行股份有限公司": "是",
    "宁波银行股份有限公司": "是",
    "浙江稠州银行股份有限公司": "是",
    "浙江泰隆商业银行股份有限公司": "是",
    "浙江民泰商业银行股份有限公司": "是",
    "浙江网商银行股份有限公司": "是",
    "宁波东海银行股份有限公司": "是",
    "宁波通商银行股份有限公司": "是"
}

COL_MUST_HAVE_SHIBOR = ["日期","onshibor"]
COL_MUST_HAVE_NIHUIGOU = ["合同号","交易号","外部成交编号","投组中文名","买卖方向","交易对手","起息日","到期日","首期结算金额","回购利率","债券名称","债券类型"]
########################################
def make_flag(a):
    if a == '国家开发银行债' or a == '记账式国债' or a == '中国农业发展银行债' or a == '中国进出口银行债':
        return 0
    else:
        return 1

def Set_Fake_Rate(a, b):
    if b == 1:
        return 100  # 设置一个比较大的值，后期需要从大到小排列，将价值连城的交易放在最首位
    else:
        return a

def Adjust_Trade_Vol(a, b, c, trade_vol):
    if c == 1:
        return (a - b + trade_vol)
    else:
        return a


def highlight_row(row):
    s = row['fuzhu']
    if s == "有背景色":
        css = 'background-color: gray'
        return [css] * len(row)
    return [""] * len(row)

class EmittingStr(QtCore.QObject):
    textWritten = QtCore.pyqtSignal(str) #定义一个发送str的信号
    def write(self, text):
         self.textWritten.emit(str(text))


########################################
class LogicWizard(QWizard,Ui_Wizard):
    def __init__(self):
        super(LogicWizard, self).__init__()

        self.setupUi(self)
        # 重定位，为了实现将控制台的信息展现在界面中
        sys.stdout = EmittingStr()
        sys.stderr = EmittingStr()
        sys.stdout.textWritten.connect(self.outputWritten)
        sys.stderr.textWritten.connect(self.outputWrittenForError)

        #sys.stdout = console_plat  # 依然返回原来的stdout设置，将print信息展现在控制台中
        #sys.stderr = console_plat  # 依然返回原来的stdout设置，将错误信息展现在控制台中

        #把默认的英文翻译成想表达的中文
        self.setButtonText(QWizard.NextButton,"下一步")
        self.setButtonText(QWizard.BackButton, "上一步")
        self.setButtonText(QWizard.CancelButton, "退出")
        self.setButtonText(QWizard.FinishButton, "运行")

        #进度条先隐藏
        self.progressBar.setVisible(False)
        self.label_6.setVisible(False)
        #self.label_7.setVisible(False)

        #cover_img = os.path.abspath(r'D:\snpythonproject\JZLC2020V1\valuezone.png')
        #image = QtGui.QPixmap(cover_img)
        image = QtGui.QPixmap("valuezone.png")
        self.label_8.setPixmap(image)
        self.label_8.setScaledContents(True)


        self.dateEdit.setDate(QtCore.QDate(2022, 1, 19))
        self.dateEdit_2.setDate(QtCore.QDate(2022, 5, 1))
        #self.lineEdit.setText(r"C:/Users/Doris/Desktop/xuqing/SHIBOR.xlsx")
        #self.lineEdit_2.setText(r"C:/Users/Doris/Desktop/xuqing/alltrades.xlsx")


        self.pushButton.clicked.connect(self.import_shibor_file)
        self.pushButton_2.clicked.connect(self.import_nihuigou_file)
        self.pushButton_3.clicked.connect(self.select_save_path)



    #用于将指令台中的信息，展示到textBrowser中
    def outputWritten(self, text):
        cursor = self.textBrowser.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.textBrowser.setTextCursor(cursor)
        self.textBrowser.ensureCursorVisible()

        # file = open("log.txt","a+",encoding="utf-8")
        # file.write(text)
        # file.close()

    # 用于将报错信息指令台中的信息，展示到同目录的text中
    def outputWrittenForError(self, text):
        # with open('out.txt', 'w+') as file:
        #     sys.stdout = file  # 标准输出重定向至文件
        #     print(text)
        file = open("Errorlog.txt","a+",encoding="utf-8")
        file.write(text)
        file.close()



    def dateCheck(self):
        if self.dateEdit_2.date()>self.dateEdit.date():
            return True
        else:

            return False

    @staticmethod
    def checkColName(df:pd.DataFrame,COL_MUST_HAVE:list):
        col_name_list =df.columns.tolist()
        a = len(set(COL_MUST_HAVE)-set(col_name_list))
        if a ==0:
            return True
        else:
            #QMessageBox.information(self,'对话框', '这是一个提醒框', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            return False

    def checkShiborDate(self,df_shibor):
        datelist = df_shibor["日期"].to_list()
        min_date = min(datelist)
        max_date = max(datelist)
        a = self.dateEdit.text()
        b = self.dateEdit_2.text()

        date_delta1 = pd.Timestamp(a) - min_date
        date_delta2 = max_date -pd.Timestamp(b)

        if date_delta1.days>7 and date_delta2.days>7:
            return True
        else:
            print("shibor的起息日和到期日需要均大于统计区间7天")
            return False



    def import_shibor_file(self):

        file = QFileDialog()  # 创建文件对话框
        file.setDirectory(r"C:\\")  # 设置初始路径为C盘
        file.setNameFilter("Excel文件(*.xls *.xlsx)")  # ！！！！！要用英语的（），否则过滤文件以后，找不到xls xlsx。
        if file.exec_():  # 判断是否选择了文件
            filename = file.selectedFiles()[0]  # 获取选择的文件
            self.lineEdit.setText(filename)  # 将选择的文件显示在文本框中
            # import os  # 导入模块


    def import_nihuigou_file(self):
        file = QFileDialog()  # 创建文件对话框
        file.setDirectory("C:/")  # 设置初始路径为C盘
        file.setNameFilter("Excel文件(*.xls *.xlsx)")  # ！！！！！要用英语的（），否则过滤文件以后，找不到xls xlsx。
        if file.exec_():  # 判断是否选择了文件
            filename = file.selectedFiles()[0]  # 获取选择的文件
            self.lineEdit_2.setText(filename)  # 将选择的文件显示在文本框中


    def select_save_path(self):
        self.dir = QFileDialog.getExistingDirectory(None, "选取文件夹","C:/")  # getExistingDirectory是一个静态函数staticmethod所以不需要实例化  查看“https://www.riverbankcomputing.com/static/Docs/PyQt5/api/qtwidgets/qfiledialog.html”
        self.lineEdit_4.setText(self.dir)

    @staticmethod
    def Holiday_ONSHIBOR_fillna(shibor_file_path,startdate,enddate,save_path):

        date_index = pd.date_range(startdate,enddate)
        s1 = pd.Series(date_index)

        df = pd.read_excel(shibor_file_path)
        df = df.set_index("日期")
        s2 = s1.map(df.to_dict()["onshibor"])
        df1 = pd.DataFrame({"日期": s1, "onshibor": s2})
        df1["onshibor"] = df1["onshibor"].fillna(method='ffill')

        save_path = os.path.join(save_path, "[1]每日的SHIBOR（假日的SHIBOR等同于前一日）.xlsx")

        df1.to_excel(save_path, index=False)
        print("完成>>>>[1]每日的SHIBOR（假日的SHIBOR等同于前一日）.xlsx")
        return df1

    @staticmethod
    def Generate_WorkdayBeforeFirstHoliday_And_Intervaldays(shibor_file_path,save_path):
        """
         [2]假日第一天的前一工作日及该工作日距离下一个工作日的天数
        :return:
        """
        df = pd.read_excel(shibor_file_path)
        df["日期"] = df["日期"].astype(str)
        df["日期"] = pd.to_datetime(df["日期"])

        df["日期向上错位移一行"] = df["日期"].shift(-1)
        df["间隔天数"] = df["日期向上错位移一行"] - df["日期"]
        df["间隔天数"] = df["间隔天数"].astype('timedelta64[D]')
        df = df[df["间隔天数"] > 1]
        df = df[["日期", "间隔天数"]]
        save_path = os.path.join(save_path, "[2]假日第一天的前一工作日及该工作日距离下一个工作日的天数.xlsx")

        df.to_excel(save_path, index=False)
        print("完成>>>>[2]假日第一天的前一工作日及该工作日距离下一个工作日的天数.xlsx")
        return df

    @staticmethod
    def select_nihuigou_with_lilv_bond(trades_file_path,save_path):
        df = pd.read_excel(trades_file_path)
        # 生成一个flag1列，如果债券类型为国债、政策性银行债，则flag1标注为0，否则标注为1
        df['flag1'] = df.apply(lambda x: make_flag(x['债券类型']), axis=1)
        df1 = df.groupby(by='交易号').sum()['flag1']
        my_dict = df1.to_dict()
        # 将df1转化为dict后和原数据进行MAP操作，key为[外部成交编号]
        df['flag2'] = df["交易号"].map(my_dict)
        # 选取flag2为0的所有交易（即表示这些交易质押的都是利率债）
        df = df[df["flag2"] == 0]

        df["起息日"] = pd.to_datetime(df["起息日"])
        df["到期日"] = pd.to_datetime(df["到期日"])
        df['间隔天数'] = df["到期日"] - df["起息日"]
        df["间隔天数"] = df["间隔天数"].astype('timedelta64[D]')
        df["首期结算金额"] = -df["首期结算金额"]
        df = df[["合同号", "交易号", "外部成交编号", "投组中文名", "买卖方向", "交易对手", "起息日", "到期日", "首期结算金额", "回购利率", "债券名称", "债券类型", "间隔天数"]]

        save_path = os.path.join(save_path, "[3]所有质押利率债的逆回购交易.xlsx")

        df.to_excel(save_path, index=False)
        print("完成>>>>[3]所有质押利率债的逆回购交易.xlsx")
        return df
        """
        :return:
        [4]假日第一天的前一工作日符合要求的交易
        """

    @staticmethod
    def Select_WorkdayBeforeFirstHoliday_Trades(df_holiday,df_trades,save_path):
        # df_holiday = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[2]假日第一天的前一工作日及该工作日距离下一个工作日的天数.xlsx")
        # df_trades = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[3]所有质押利率债的逆回购交易.xlsx")

        df_holiday.rename(columns={'日期': '起息日'}, inplace=True)


        df = pd.merge(df_trades, df_holiday, on=['起息日', '间隔天数'], how="inner")
        save_path = os.path.join(save_path, "[4]假日第一天的前一工作日符合要求的交易.xlsx")
        df.to_excel(save_path, index=False)
        print("完成>>>>[4]假日第一天的前一工作日符合要求的交易.xlsx")
        return df

    @staticmethod
    def Select_JiaZhiLianCheng_Shibor_Rate_Trades(df_onenight_shibor,df_trades,save_path):
        #df_onenight_shibor = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[1]每日的SHIBOR（假日的SHIBOR等同于前一日）.xlsx")

        df_onenight_shibor = df_onenight_shibor.rename(columns={'日期': '起息日'})
        #df_trades = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[3]所有质押利率债的逆回购交易.xlsx")

        df1 = pd.merge(df_trades, df_onenight_shibor, on='起息日', how="left")
        df1['成交价格与O/NSHIBOR的价差'] = df1['回购利率'] - df1['onshibor']
        df1['是否为价值连城客户'] = df1["交易对手"].map(MEMBER_MAP_DICT)
        df1 = df1[(df1['是否为价值连城客户'] == '价值连城') & (df1['成交价格与O/NSHIBOR的价差'] == 0)]
        df1['优先统计'] = 1
        save_path = os.path.join(save_path, "[5]交易对手为价值连城成员且回购利率等于SHIBOR隔夜利率的交易.xlsx")
        df1.to_excel(save_path, index=False)

        print("完成>>>>[5]交易对手为价值连城成员且回购利率等于SHIBOR隔夜利率的交易.xlsx")
        return df1

    @staticmethod
    def Select_One_Intervalday_Trades(df_trades,save_path):
        """
        #[6]起息日与到期日间隔1天的所有逆回购交易
        :return:
        """
        #df_trades = pd.read_excel("[3]所有质押利率债的逆回购交易.xlsx")
        df = df_trades[df_trades['间隔天数'] == 1]
        save_path = os.path.join(save_path, "[6]起息日与到期日间隔1天的所有逆回购交易.xlsx")

        df.to_excel(save_path, index=False)
        print("完成>>>>[6]起息日与到期日间隔1天的所有逆回购交易.xlsx")
        return df


    @staticmethod
    def Generate_Wanted_Trade_Pool(df4,df5,df6,save_path):
        """
        #[7]所有符合要求待选的交易
        :return:
        """
        # df6 = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[6]起息日与到期日间隔1天的所有逆回购交易.xlsx")
        # df4 = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[4]假日第一天的前一工作日符合要求的交易.xlsx")
        # df5 = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[5]交易对手为价值连城成员且回购利率等于SHIBOR隔夜利率的交易.xlsx")
        df7 = pd.concat([df4, df5, df6], ignore_index=True)
        df7["优先统计"] = df7["优先统计"].fillna(0)
        df7["是否为价值连城客户"] = df7["是否为价值连城客户"].fillna("否")
        df7 = df7.sort_values(by='优先统计', ascending=False)
        df7 = df7.drop_duplicates(subset='交易号', keep='first')
        save_path = os.path.join(save_path, "[7]所有符合要求待选的交易.xlsx")

        df7.to_excel(save_path, index=False)
        print("完成>>>>[7]所有符合要求待选的交易.xlsx")
        return df7

    @staticmethod
    def CopyAndInsert_N_Intervalday_Trades(df7,save_path):
        """
        # 该函数实现：如果实际占款天数为N天，则自动复制并插入N-1条相同的记录,并将复制的交易的起息日往后加一天，其他要素不变
        # [8]所有符合要求待选的交易_复制占用天数大于1的交易（相同交易起息日+1天）.xlsx
        :return:
        """
        #df7 = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[7]所有符合要求待选的交易.xlsx")
        # 产生一个新的DF，根据间隔天数的多少，重复交易号，该DF用于和原来的DF进行来连接，最终目的是若该笔交易的间隔天数N，则产生N条该笔记录
        tradeid = df7["交易号"].tolist()
        cishu = df7["间隔天数"].tolist()

        new_tradeid_list = []
        qixiri_delta_list = []
        for i in range(len(tradeid)):
            if cishu[i] == 1:
                new_tradeid_list.append(tradeid[i])
                qixiri_delta_list.append(0)
            else:
                for k in range(int(cishu[i])):
                    new_tradeid_list.append(tradeid[i])
                    qixiri_delta_list.append(k)
        data1 = {"交易号": new_tradeid_list, "起息日delta": qixiri_delta_list}
        df = pd.DataFrame(data1)

        df8 = pd.merge(df, df7, on=["交易号"], how="right")
        # 原本df2["起息日delta"]是 int64类型，该函数将int转换为timedela，注；unit = "D"意思是将int N,转换为N days
        df8["起息日delta"] = pd.to_timedelta(df8["起息日delta"], unit="D")
        df8["起息日（间隔日大于1日的有调整）"] = df8["起息日"] + df8["起息日delta"]
        save_path = os.path.join(save_path, "[8]所有符合要求待选的交易_复制占用天数大于1的交易（相同交易起息日+1天）.xlsx")
        df8.to_excel(save_path, index=False)
        print("完成>>>>[8]所有符合要求待选的交易_复制占用天数大于1的交易（相同交易起息日+1天）.xlsx")
        return df8



    def Insert_Virtual_Trades_HZbank(self,df,df8,save_path):
        """
        插入交易对手为杭州银行的虚拟逆回购交易，
        目的：如果当日所有符合要求的质押式逆回购交易数量不到TRADE_VOL，那智能用虚拟交易来充数
        金额等于价值连城拆借规模上限 = TRADE_VOL（极端假设当天完全没有符合要求的真实的质押式逆回购交易）;
        期限=1天;
        回购利率=当日O/N SHIBOR
        :return:
        """
        #df = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[1]每日的SHIBOR（假日的SHIBOR等同于前一日）.xlsx")
        df["交易对手"] = "杭州银行股份有限公司"
        df["回购利率"] = df["onshibor"]
        df["债券名称"] = "虚拟利率债"
        df["债券类型"] = "虚拟利率债"
        df["间隔天数"] = 1
        df["优先统计"] = 0
        df["买卖方向"] = "逆回购"
        df["投组中文名"] = "虚拟插入的交易"
        df["外部成交编号"] = "虚拟插入的交易"
        df["是否为价值连城客户"] = "否"

        df["首期结算金额"] = int(self.lineEdit_3.text())*100000000
        df = df.rename(columns={'日期': '起息日（间隔日大于1日的有调整）'})
        df["起息日"] = df['起息日（间隔日大于1日的有调整）']
        df["起息日delta"] = 1
        df["起息日delta"] = pd.to_timedelta(df["起息日delta"], unit="D")
        df["到期日"] = df["起息日（间隔日大于1日的有调整）"] + df["起息日delta"]

        #df8 = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[8]所有符合要求待选的交易_复制占用天数大于1的交易（相同交易起息日+1天）.xlsx")
        df9 = pd.concat([df8, df])
        save_path = os.path.join(save_path, "[9]所有符合要求待选的交易_插入杭州银行虚拟交易.xlsx")
        df9.to_excel(save_path, index=False)
        print("完成>>>>[9]所有符合要求待选的交易_插入杭州银行虚拟交易.xlsx")
        return df9

    @staticmethod
    def Sort_Trades_JiaZhiLianCheng_Priority(df9,save_path):
        #df9 = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[9]所有符合要求待选的交易_插入杭州银行虚拟交易.xlsx")
        # 如果['优先统计']是1，则设置一个假的利率100.为了以后根据fake_rate排序能排在前面
        df9['fake_rate'] = df9.apply(lambda x: Set_Fake_Rate(x['回购利率'], x['优先统计']), axis=1)
        # 以起息日为单位，将“回购利率”按照从大到小排列
        df10 = df9.sort_values(by=['起息日', 'fake_rate'], ascending=[True, False])
        save_path = os.path.join(save_path, "[10]所有符合要求待选的交易_已按交易日期及回购利率排序_价值连城且回购利率等于SHIBOR隔夜的交易优先.xlsx")
        df10.to_excel(save_path, index=False)
        print("完成>>>>[10]所有符合要求待选的交易_已按交易日期及回购利率排序_价值连城且回购利率等于SHIBOR隔夜的交易优先.xlsx")
        return df10

    def Select_Trades_By_Cumsum(self,df10,save_path):
        '''

        :return: 返回并保存最终符合要求的交易。
                会存在这样一种情况：若一笔价值连城成员的交易，交易日期为2019-06-29，占款天数为7天，该函数操作之前该笔交易会产生7另外6笔交易
                起息日分别为2019-06-29、2019-06-30、2019-07-01、2019-07-02、2019-07-03、2019-07-04、2019-07-05

                符合要求的交易的选取步骤如下：
                按照每天为单位进行统计
                步骤1、价值连城客户的交易优先纳入统计范围（由函数Sort_Trades_JiaZhiLianCheng_Priority进行处理）
                步骤2、回购品种为隔夜（如果是假日前一个工作的交易，则取类隔夜交易）的交易按照隔夜利率从高到底排序
                步骤3、步骤1+步骤2的交易量汇总，直至累加至14亿元。选取并保存累加至14亿元的所有交易
        '''
        pool_vol = int(self.lineEdit_3.text())*100000000
        #df10 = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[10]所有符合要求待选的交易_已按交易日期及回购利率排序_价值连城且回购利率等于SHIBOR隔夜的交易优先.xlsx")
        df10['累计加总额度'] = df10.groupby('起息日（间隔日大于1日的有调整）')['首期结算金额'].cumsum()
        df10['是否大于阀值'] = df10['累计加总额度'].apply(lambda x: 0 if x < pool_vol else 1)  # 判断累计加总之和是否大于14亿元
        df10['是否大于阀值的累计和'] = df10.groupby('起息日（间隔日大于1日的有调整）')['是否大于阀值'].cumsum()
        df10['修正后的交易量'] = df10['首期结算金额']

        df10['修正后的交易量'] = df10.apply(lambda x: Adjust_Trade_Vol(x['首期结算金额'], x['累计加总额度'], x['是否大于阀值的累计和'], pool_vol),
                                     axis=1)
        df11 = df10[(df10['是否大于阀值的累计和'] < 2)]
        save_path = os.path.join(save_path, "[11]最终被选出的交易.xlsx")
        df11.to_excel(save_path, index=False)
        print("完成>>>>[11]最终被选出的交易.xlsx")
        return df11

    def Calc_Interest(self, df11,save_path):
        pool_vol = int(self.lineEdit_3.text()) * 100000000
        #df11 = pd.read_excel(r"D:\snpythonproject\JZLC2020V1\[11]最终被选出的交易.xlsx")
        df11['当日收益'] = df11['修正后的交易量'] * df11['回购利率'] * 1 / 36500
        #df11.loc[:,'当日收益'] = df11[:,'修正后的交易量'] * df11[:,'回购利率'] * 1 / 36500
        total_interest = df11['当日收益'].sum()
        a = pd.to_datetime(self.dateEdit.text())
        b = pd.to_datetime(self.dateEdit_2.text())

        interval_days = b - a  # 此时interval_days是Timedelta类型
        interval_days = interval_days.days  # 提取Timedelta类型中的天数，返回的是int型

        # 计算价值连城运作的年化收益率
        profit_interest_rate = total_interest / pool_vol / interval_days * 365 * 100
        profit_interest_rate = round(profit_interest_rate, 2)
        total_interest = round(total_interest, 2)
        total_interest = format(total_interest, ',')  # 添加千分位


        #该段代码用于将结果明细通过背景色不同展示出来
        datelist = df11["起息日（间隔日大于1日的有调整）"].tolist()
        datelist = list(set(datelist))
        datelist = sorted(datelist)

        fuzhu_list = []
        for i in range(len(datelist)):
            if i % 2 == 1:
                fuzhu_list.append("无背景色")
            else:
                fuzhu_list.append("有背景色")

        fuzhu_dict = dict(zip(datelist, fuzhu_list))
        df11["fuzhu"] = df11["起息日（间隔日大于1日的有调整）"].map(fuzhu_dict)

        df11.sort_values("起息日（间隔日大于1日的有调整）", inplace=True, ascending=True)

        df11 = df11.reset_index()

        df12 = df11[["交易号", "外部成交编号", "交易对手", "是否为价值连城客户", "买卖方向", "回购利率", "首期结算金额", "起息日", "到期日", "起息日（间隔日大于1日的有调整）",
                      "修正后的交易量", "当日收益","fuzhu"]]

        df12 = df12.style.apply(highlight_row, axis=1)

        save_path = os.path.join(save_path, "[12]最终被选出的交易_计算结果.xlsx")

        writer = pd.ExcelWriter(save_path)
        df12.to_excel(writer, sheet_name='最终结果')
        writer.save()
        print("完成>>>>[12]最终被选出的交易_计算结果.xlsx")

        #
        # df12.to_excel(save_path, engine = 'openpyxl')
        # print("完成>>>>[12]最终被选出的交易_计算结果.xlsx")
        return total_interest,profit_interest_rate


    def run(self):
        shibor_file_path = self.lineEdit.text()
        trades_file_path = self.lineEdit_2.text()
        file_save_path =self.lineEdit_4.text()
        df_onenight_shibor = self.Holiday_ONSHIBOR_fillna(shibor_file_path,self.dateEdit.text(),self.dateEdit_2.text(),file_save_path)
        self.progressBar.setValue(20)
        df_holiday = self.Generate_WorkdayBeforeFirstHoliday_And_Intervaldays(shibor_file_path,file_save_path)
        self.progressBar.setValue(25)
        df_trades = self.select_nihuigou_with_lilv_bond(trades_file_path,file_save_path)
        self.progressBar.setValue(30)
        df4 = self.Select_WorkdayBeforeFirstHoliday_Trades(df_holiday,df_trades,file_save_path)
        self.progressBar.setValue(35)
        df5 = self.Select_JiaZhiLianCheng_Shibor_Rate_Trades(df_onenight_shibor,df_trades,file_save_path)
        self.progressBar.setValue(40)
        df6 = self.Select_One_Intervalday_Trades(df_trades,file_save_path)
        self.progressBar.setValue(45)
        df7 = self.Generate_Wanted_Trade_Pool(df4,df5,df6,file_save_path)
        self.progressBar.setValue(50)
        df8 = self.CopyAndInsert_N_Intervalday_Trades(df7,file_save_path)
        self.progressBar.setValue(60)
        df9 = self.Insert_Virtual_Trades_HZbank(df_onenight_shibor,df8,file_save_path)
        self.progressBar.setValue(70)
        df10 = self.Sort_Trades_JiaZhiLianCheng_Priority(df9,file_save_path)
        self.progressBar.setValue(80)
        df11 = self.Select_Trades_By_Cumsum(df10,file_save_path)
        self.progressBar.setValue(90)
        profit_result = self.Calc_Interest(df11,file_save_path)
        self.progressBar.setValue(100)
        return profit_result



    #在点击“next”或者“”finish“时，都会调用validateCurrentPage，
    # 当validateCurrentPage返回的值为False时，页面停留，当返回值为True时，进入下一页
    def validateCurrentPage(self):
        try:
            if self.currentId() ==0:
                if self.lineEdit.text() =="":
                    QMessageBox.critical(self,'输入报错','请导入SHIBOR数据',QMessageBox.Ok)
                    return False

                elif self.dateCheck() == False:
                    QMessageBox.critical(self, '输入报错', '统计到期日应大于统计起始日', QMessageBox.Ok)
                    return False
                else:
                    df_shibor = pd.read_excel(self.lineEdit.text())
                    if self.checkColName(df_shibor,COL_MUST_HAVE_SHIBOR) ==False:
                        QMessageBox.critical(self, '输入报错', "导入的SHIBOR文件应包含{}当中的所有字段".format(COL_MUST_HAVE_SHIBOR), QMessageBox.Ok)
                        return False
                    elif self.checkShiborDate(df_shibor) == False:
                        QMessageBox.critical(self, '对话框', '导入的SHIBOR文件前后日期应比统计的前后日期向前后延长7天',QMessageBox.Ok)
                    else:
                        return True
            else:
                self.textBrowser.clear()
                self.label_7.setText("统计结果")
                if self.lineEdit_2.text() == "":
                    QMessageBox.critical(self, '输入报错', '请导入逆回购数据', QMessageBox.Ok)
                    return False
                elif self.lineEdit_4.text() =="":
                    QMessageBox.critical(self, '输入报错', '请选择文件的保存路径', QMessageBox.Ok)
                    return False
                else:
                    # 显示进度条
                    self.label_6.setVisible(True)
                    self.progressBar.setVisible(True)
                    self.progressBar.setValue(5)
                    df_nihugou = pd.read_excel(self.lineEdit_2.text())
                    self.progressBar.setValue(10)
                    if self.checkColName(df_nihugou, COL_MUST_HAVE_NIHUIGOU) == False:
                        QMessageBox.critical(self, '输入报错', '导入的逆回购文件应包含{}当中的所有字段'.format(COL_MUST_HAVE_NIHUIGOU), QMessageBox.Ok)
                        return False
                    else:
                        try:
                            profit_result = self.run()
                            QMessageBox.information(self, '信息提示', '报告老板，计算完成', QMessageBox.Ok)
                            self.label_7.setVisible(True)
                            self.label_7.setText("价值连城资金池在{}至{}期间\n的利息总收入为{}元，\n年化收益率为{}%".format(self.dateEdit.text(),self.dateEdit_2.text(),profit_result[0],profit_result[1]))
                            return False
                        except:
                            traceback.print_exc()
                            QMessageBox.critical(self, '输入报错', '运行报错，请查看ERROR_LOG', QMessageBox.Ok)
                            return False
        except:
            traceback.print_exc()
            QMessageBox.critical(self, '输入报错', '运行报错，请查看ERROR_LOG', QMessageBox.Ok)
            return False


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    ui = LogicWizard()
    ui.show()
    sys.exit(app.exec_())