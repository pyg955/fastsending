from PySide2.QtWidgets import QApplication, QMainWindow, QTabWidget, QListWidget, QListWidgetItem, QLabel,QSpinBox,QMessageBox,QDesktopWidget
from PySide2.QtUiTools import QUiLoader
import win32com.client
from jinja2 import Environment, FileSystemLoader
import win32timezone
import logging


# 设置日志格式
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# 创建一个FileHandler来输出到文件
file_handler = logging.FileHandler('error.log')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)
logging.getLogger('').addHandler(file_handler)

clicked_messages = 0
clicked_message = 0

# outlook 相关方法
class Outlook:
    
    def __init__(self):
        self.Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.root_folder = self.Outlook.Folders.Item("yigang@mgtv.com")
        self.env = Environment(loader=FileSystemLoader('.'))  # 邮件模板

    def get_root_folder(self):
        return self.root_folder

    def get_folder_emails(self,root_folder,folder_name):
    # 传入邮件中文件夹实例、文件夹名称，放对应文件夹下所有邮件列表实例
        sub_folders = root_folder.Folders
        for sub_folder in sub_folders:
            if sub_folder.Name == folder_name:
                print("正在读取文件夹:", sub_folder.Name)
                self.messages = sub_folder.Items
                self.messages.Sort("[ReceivedTime]", True)  # 按接收时间倒序
        return self.messages

    def get_email(self,messages,email_name):
    # 传入邮件列表实例、邮件名称，返回具体邮件实例
        for messsage in messages:
            if messsage.Subject == email_name:
                return messsage

    def get_emails_list(self,messages,num):
        """
        返回邮件列表
        使用一个for循环遍历邮件列表，并使用enumerate()函数获取每个邮件对象的索引和值。
        在循环中，检查索引是否等于指定的num值。如果是，则打印该邮件的主题并退出循环
        """
        temp_list = []
        for index in range(len(messages)):
            if index == num:
                break
            print("邮件主题：", index+1, messages[index].Subject)
            temp_list.append(messages[index].Subject)
        return temp_list

    def get_email_info(self,message,email_info):
    # 传入邮件实例、邮件信息标识，返回对于标识下的信息
        if email_info == "Subject":
            print("邮件主题:", message.Subject)
            return message.Subject
        elif email_info == "SenderName":
            print("发件人:", message.SenderName)
            return message.SenderName
        elif email_info == "CC":
            print("抄送人:", message.CC)
            return message.CC
        elif email_info == "SentOn":
            print("时间:", message.SentOn)
            return message.SentOn
        elif email_info =="Body":
            print("内容:", message.Body)
            return message.Body
        elif email_info == "HTMLBody":
            print("内容:", message.HTMLBody)
            return message.HTMLBody

    

# 通用相关方法
class Common:
    

    def __init__(self,ui,outlook):
        self.ui = ui
        self.Outlook = outlook

    def get_current_tab_info(self):
        # 获取当前tab的index和text
        current_index = self.ui.tabWidget.currentIndex() # 获取当前选项卡的索引
        current_text = self.ui.tabWidget.tabText(current_index)  # 获取当前选项卡的文本内容
        print(current_text)
        return current_index,current_text

    def get_readnum_value(self):
        return self.ui.spinBox.value()

    def get_clicked_title(self,item):
        # 根据选中的邮件获取标题
        return item.text()

    def get_clicked_message(self,item):
        # 获取当前选中的邮件实例
        global clicked_messages,clicked_message
        _,current_text = self.get_current_tab_info()
        messages = self.Outlook.get_folder_emails(self.Outlook.get_root_folder(),current_text)
        message = self.Outlook.get_email(messages,self.get_clicked_title(item))
        clicked_messages = messages
        clicked_message = message
        return messages,message

    def get_cc_info(self,message):
        # 获取抄送人信息
        return self.Outlook.get_email_info(message,"CC")

    def get_sender_info(self,message):
        # 获取发件人信息
        return self.Outlook.get_email_info(message,"SenderName")

    def get_senton_info(self,message):
        # 获取发件事件
        return self.Outlook.get_email_info(message, "SentOn")

    def get_email_title_edit(self):
        # 获取邮件标题编辑框中的内容
        return self.ui.lineEdit.text()

    def get_project_name_edit(self):
        # 获取项目名称编辑框中的内容
        return self.ui.lineEdit_4.text()

    def get_test_round_spinBox(self):
        # 获取测试轮次编辑框中的内容
        return self.ui.spinBox_5.value()

    def get_new_isues_spinBox(self):
        # 获取新增bug编辑框中的内容
        return self.ui.spinBox_2.value()

    def get_closed_issues_spinBox(self):
        # 获取关闭bug编辑框中的内容
        return self.ui.spinBox_3.value()

    def get_reopened_issues_spinBox(self):
        # 获取重开bug编辑框中的内容
        return self.ui.spinBox_4.value()

    def get_conclusion_cBox(self):
        # 获取测试结果编辑框中的内容
        return self.ui.comboBox.currentText()

    def get_test_env_cBox(self):
        # 获取测试结果编辑框中的内容
        return self.ui.comboBox_2.currentText()

    def get_defect_list_edit(self):
        # 获取缺陷清单编辑框中的内容
        return self.ui.lineEdit_10.text()

    def get_test_content_edit(self):
        # 获取需求清单编辑框中的内容
        return self.ui.lineEdit_11.text()

    def get_test_device_edit(self):
        # 获取设备编辑框中的内容
        return self.ui.lineEdit_13.text()

    def get_test_case_count_spinBox(self):
        # 获取测试用例数量
        return self.ui.spinBox_6.value()

    def get_auto_test_case_count_spinBox(self):
        # 获取自动划测试用例数量
        return self.ui.spinBox_7.value()

    def get_system_rediobtn(self):
        # 获取选择的系统
        android_sys = self.ui.checkBox.isChecked()  # 获取安卓等选择状态 bool型
        ios_sys = self.ui.checkBox_2.isChecked()
        win_sys = self.ui.checkBox_3.isChecked()
        hongmeng_sys = self.ui.checkBox_4.isChecked()
        mac_sys = self.ui.checkBox_5.isChecked()
        temp_dic = {"安卓":android_sys,"IOS":ios_sys,"鸿蒙":hongmeng_sys,"Windows":win_sys,"Mac":mac_sys}  # 创建字典
        temp_list = []
        for key, value in temp_dic.items():
            if value:
                temp_list.append(key)  # 如果value=true 则将key加入list，表示被选中
        result = "、".join(temp_list)
        print(result)
        return result

    def get_cc_edit(self):
        # 获取抄送人编辑框中的内容
        return self.ui.lineEdit_3.text()

    def set_title(self,text):
        # 设置标题信息
        self.ui.lineEdit.setText(text)

    def set_cc(self,text):
        # 设置抄送人
        self.ui.lineEdit_3.setText(text)

    def set_sender(self,text):
        # 设置发件人
        self.ui.lineEdit_2.setText(text)

    def reply_email(self,message,html):
        # 回复邮件
        message.Subject = f"{self.get_email_title_edit()} 第{self.get_test_round_spinBox()}轮"
        message.CC = self.get_cc_edit()
        reply = message.ReplyAll()
        reply_html = html
        newline = "<br><br\>"  # 换行符号
        line = "<hr>"  # outlook自带的分割线
        senton = self.get_senton_info(clicked_message)

        reply.HTMLBody = f"{reply_html}{newline}{line}{senton}{message.HTMLBody}"
        reply.Display()  # 用于在屏幕上展示邮件

class Stats:

    def __init__(self):
        # 从文件中加载UI定义

        # 从 UI 定义中动态 创建一个相应的窗口对象
        # 注意：里面的控件对象也成为窗口对象的属性了
        # 比如 self.ui.button , self.ui.textEdit
        self.ui = QUiLoader().load('mainUI.ui') 
        self.Outlook = Outlook()  # 实例化 Outlook类
        self.Common = Common(self.ui,self.Outlook)  # 实例化 Common类
        self.ui.comboBox.addItems(["不通过","测试通过","部分测试通过"])  # 添加测试结果选项
        self.ui.comboBox_2.addItems(["DNS","线上","DNS+线上"])  # 添加测试环境选项
        self.ui.pushButton_2.clicked.connect(self.refresh_btn)  # 刷新按钮
        self.ui.listWidget.itemClicked.connect(self.on_item_clicked)  # 列表1点击事件
        self.ui.listWidget_2.itemClicked.connect(self.on_item_clicked)  # 列表2点击事件
        self.ui.pushButton_3.clicked.connect(self.preview)   # 预览按钮点击事件
        self.ui.buttonGroup_2.buttonClicked.connect(self.Common.get_system_rediobtn)  # 系统选择事件
        self.env = Environment(loader=FileSystemLoader('.'))   # 实例化邮件模板相关内容
        self.refresh_btn()  # 启动则获取一次邮件

        """
         生成异常弹窗
        """
        self.ui.error_window = QMainWindow()
        self.ui.error_window.resize(1000, 800)
        # 获得屏幕坐标系
        self.ui.screen = QDesktopWidget().screenGeometry()

        # 计算窗口左上角的坐标
        x = (self.ui.screen.width() - self.ui.error_window.width()) / 2
        y = (self.ui.screen.height() - self.ui.error_window.height()) / 2

        # 将窗口移动到左上角的坐标
        self.ui.error_window.move(x, y)
        self.ui.error_window.setWindowTitle('异常情况')



    # 刷新按钮
    def refresh_btn(self):
        current_index,current_text = self.Common.get_current_tab_info()  # 获取当前tab的index和text
        edit_read_number = self.Common.get_readnum_value()  # 获取读数量的值
        all_emails = self.Outlook.get_folder_emails(self.Outlook.get_root_folder(),current_text) 
        email_list = self.Outlook.get_emails_list(all_emails,edit_read_number)

        self.ui.listWidget.clear()  # 清空所有的内容
        self.ui.listWidget_2.clear()
        for email in email_list:
            list_item = QListWidgetItem(email)
            if current_index == 0 :  # 判断是第一个tab
                self.ui.listWidget.addItem(list_item)
            elif current_index == 1 :  # 判断是第二个tab
                self.ui.listWidget_2.addItem(list_item)

    # 点击列表中的项
    def on_item_clicked(self,item):
        global clicked_message
        print("你点击了：" + item.text())
        self.Common.set_title("【测试结果】" + str(item.text()))  # 设置【测试结果】+标题
        self.Common.get_clicked_message(item)
        cc = self.Common.get_cc_info(clicked_message)   # 获取当前点击邮件的抄送人
        sender = self.Common.get_sender_info(clicked_message)   # 获取当前点击邮件的抄送人
        self.Common.set_cc(cc)  # 设置抄送人
        self.Common.set_sender(sender)  # 设置发件人

    # 预览按钮
    def preview(self):
        project_name = self.Common.get_project_name_edit()
        test_round = self.Common.get_test_round_spinBox()
        new_issues = self.Common.get_new_isues_spinBox()
        closed_issues = self.Common.get_closed_issues_spinBox()
        reopened_issues = self.Common.get_reopened_issues_spinBox()
        conclusion = self.Common.get_conclusion_cBox()
        defect_list = self.Common.get_defect_list_edit()
        test_content = self.Common.get_test_content_edit()
        test_env = self.Common.get_test_env_cBox()
        system = self.Common.get_system_rediobtn()
        test_device = self.Common.get_test_device_edit()
        test_case_count = self.Common.get_test_case_count_spinBox()
        template = self.env.get_template('template1.html')
        auto_test_case_count = self.Common.get_auto_test_case_count_spinBox()
        html = template.render(project_name=project_name,
                               test_round=test_round,
                               new_issues=new_issues,
                               closed_issues=closed_issues,
                               reopened_issues=reopened_issues,
                               conclusion=conclusion,
                               defect_list=defect_list,
                               test_content=test_content,
                               test_env=test_env,
                               system=system,
                               test_device=test_device,
                               test_case_count=test_case_count,
                               auto_test_case_count=auto_test_case_count )
        try:
            self.Common.reply_email(clicked_message,html)  # 回复邮件
        except:
            logging.exception("发生错误！")
            QMessageBox.critical(self.ui.error_window,
                '异常情况',
                "请先选择邮件！"
                )


if __name__ == '__main__':
    # 初始化应用程序
    app = QApplication([])
    window = Stats()
    window.ui.show()
    app.exec_()     # 运行