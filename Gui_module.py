import sys
import os

import threading

from PyQt5.QtWidgets import QSpacerItem, QFrame, QSizePolicy, QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QSplitter, QListWidget, QFileDialog, QLabel, QComboBox, QListWidgetItem
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSettings, QMutex, QObject, pyqtSignal, QTimer

from LD_dnconsole import Dnconsole
from script_module import Script_action
from Test_match_module import test_match


class ScriptExecutor(QObject):
    finished = pyqtSignal()
    progress_updated = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.running = True
        self.mutex = QMutex() #使用 QMutex 来保护共享数据的访问

        self.dnconsole = None
        self.selected_folder = None
        self.excel_list = None
        self.Working_label = None
        self.binding_handle = None

    def set_running(self, value):
        self.mutex.lock()
        self.running = value
        self.mutex.unlock()

    def get_running(self):
        self.mutex.lock()
        running = self.running
        self.mutex.unlock()
        return running

    def execute_script_task(self):
        self.set_running(True)
        # 获取LD子窗口句柄，用于执行脚本按钮
        self.binding_handle = self.parent().get_binding_handle()
        # 窗口置顶，否则键盘命令“esc”按键会出错 》》》esc按钮被占用，建议换按钮
        # win32gui.SetWindowPos(self.binding_handle, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        # win32gui.SetForegroundWindow(self.binding_handle)

        excell_path_list = [self.excel_list.item(i).text() for i in range(self.excel_list.count())]

        if not excell_path_list:
            self.progress_updated.emit("没有需要处理的Excel文件")
            return
        self.progress_updated.emit("执行脚本")
        # 检查并初始化 self.dnconsole
        if self.dnconsole is None:
            if self.selected_folder is None:
                self.progress_updated.emit("请先选择雷电安装文件夹")
                return
            self.parent().start_program_and_app()

        for excell_path in excell_path_list:
            if not self.get_running():
                break
            try:
                Script_action.execute_script_from_excel(self.dnconsole, self.binding_handle, excell_path, self)
            except Exception as e:
                self.progress_updated.emit(f"执行脚本时发生错误: {str(e)}")
                print(f"Error executing script from {excell_path}: {e}")
                # 可以选择在这里停止执行，或者继续尝试下一个文件
                continue

        self.finished.emit()  # 发出完成信号


def resource_path(relative_path):
    #这个方法通过检查 sys._MEIPASS 是否存在来判断程序是否已经打包，并据此构建正确的文件路径。
    """ 获取资源的绝对路径 """
    try:
        # PyInstaller 创建临时文件夹，_MEIPASS 是其中的一个特殊变量
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("MyCompany", "MyApp")
        self.initUI()

        #使用 QTimer 定期检查任务状态并更新 UI
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_ui)
        self.timer.start(1000)  # 每秒更新一次

        self.running = True
        self.dnconsole = None  # 初始化dnconsole属性
        self.top_level_handle = None # 存储雷电模拟器主窗口的句柄（顶层句柄）
        self.binding_handle = None # 存储雷电模拟器子窗口的句柄（绑定句柄）

        self.folder_for_test_image = ""  # 初始化一个属性来存储上次打开的文件夹路径
        self.folder_for_test_template = ""  # 初始化一个属性来存储上次打开的文件夹路径

        self.selected_folder = None  # 存储选择的文件夹路径
        self.selected_app = None  # 存储选中的App名称

        self.test_image_full_path = None # 存储选择的测试大底图
        self.template_folder = None # 存储选择的测试模板图片文件夹

        # self.excel_list = QListWidget(self)  # 存储Excel文件列表

        self.script_executor = ScriptExecutor(self)  # 创建脚本执行器实例
        self.script_executor.progress_updated.connect(self.update_working_label)
        self.script_executor.finished.connect(self.on_script_finished) # 连接信号到槽

    def initUI(self):
        # 制作逻辑：
        # 窗口 -> 布局 -> 子窗口/子布局 -> 子布局
        # 大关系：创建按钮 -> 按钮应用到布局  -> 布局应用到显示窗口
        # 实现按钮功能

        # 设置窗口大小和标题
        self.setGeometry(300, 300, 1500, 600)
        self.setWindowTitle('LD-Easy-Job')

        # 设置窗口图标
        # self.setWindowIcon(QIcon('Game+.png'))  # 替换为你的Logo文件路径
        # 获取图标文件的路径
        icon_path = resource_path('Game+.png')
        # 设置窗口图标
        self.setWindowIcon(QIcon(icon_path))

        # 创建主布局
        main_layout = QVBoxLayout()

        # 第一行布局 - 程序启动内容
        first_row_layout = QHBoxLayout() #创建 第一行 布局

        # 雷电安装文件夹选择部分-按钮
        leidian_layout = QVBoxLayout()  # 创建 第一行局部 布局
        self.leidian_path_label = QLabel('未选择', self)  # 制作显示列表
        leidian_layout.addWidget(self.leidian_path_label)  # 添加显示列表（窗口）到布局
        self.leidian_folder_button = QPushButton('选择雷电安装文件夹', self)  # 制作按钮
        self.leidian_folder_button.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        leidian_layout.addWidget(self.leidian_folder_button)  # 添加按钮（窗口）到布局

        # App名字 - 选择列表
        app_layout = QVBoxLayout()
        spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Fixed)
        app_layout.addItem(spacer)
        self.app_combobox = QComboBox(self)
        self.app_combobox.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        app_layout.addWidget(self.app_combobox)

        # 启动程序按钮
        self.start_program_button = QPushButton('启动程序', self)
        self.start_program_button.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

        # 将两个部分放入一个水平布局
        first_row_layout.addLayout(leidian_layout) #将局部布局 添加到 第一行布局中
        first_row_layout.addLayout(app_layout)  #将局部布局 添加到 第一行布局中
        first_row_layout.addWidget(self.start_program_button)  # 将启动按钮（窗口） 添加到 第一行布局中

        # ========================================

        # 创建一个水平分割线
        h_line = QFrame()
        h_line.setFrameShape(QFrame.HLine)
        # h_line.setFrameShadow(QFrame.Sunken)  # 可选，设置阴影效果
        h_line.setLineWidth(3)  # 设置线条宽度
        # h_line.setMidLineWidth(30)  # 设置线条中间部分的宽度
        h_line.setStyleSheet("margin-top: 50px; margin-bottom: 50px;")  # 设置上下边距

        # ========================================

        # 第二行布局 - 测试组件 & 脚本组件
        second_row_layout = QHBoxLayout()
        splitter = QSplitter()

        # 第二行第一列布局 - 测试组件
        test_widget = QWidget()
        test_layout = QVBoxLayout()

        self.test_image_path_label = QLabel('未选择', self)  # 制作 文件 显示列表
        self.load_test_image_button = QPushButton('读取测试图片', self)
        self.template_folder_label = QLabel('未选择', self)    # 制作 文件 显示列表
        self.load_template_folder_button = QPushButton('读取测试模板文件夹', self)
        self.match_image_button = QPushButton('图片匹配测试', self)

        test_layout.addWidget(self.test_image_path_label)  # 添加显示列表（窗口）到布局
        test_layout.addWidget(self.load_test_image_button)
        test_layout.addWidget(self.template_folder_label)
        test_layout.addWidget(self.load_template_folder_button)
        test_layout.addWidget(self.match_image_button)

        test_widget.setLayout(test_layout) #将 测试组件的布局 添加到 测试组件窗口

        # 第二行第二列布局 - 脚本组件1
        script_widget = QWidget()
        script_layout = QVBoxLayout()

        self.excel_list = QListWidget() #制作显示列表
        self.load_excel_button = QPushButton('载入Excell文件', self)
        self.delete_excel_button = QPushButton('删除Excell文件', self)
        self.move_up_button = QPushButton('上移', self)
        self.move_down_button = QPushButton('下移', self)
        script_layout.addWidget(self.excel_list) #添加显示列表到第二列布局
        script_layout.addWidget(self.load_excel_button)
        script_layout.addWidget(self.delete_excel_button)
        script_layout.addWidget(self.move_up_button)
        script_layout.addWidget(self.move_down_button)

        script_widget.setLayout(script_layout) #将 脚本组件1的布局 添加到 脚本组件1窗口


        # 创建一个垂直分割线
        v_line = QFrame()
        v_line.setFrameShape(QFrame.VLine)
        # v_line.setFrameShadow(QFrame.Sunken)  # 可选，设置阴影效果
        v_line.setLineWidth(1)  # 设置线条宽度
        # v_line.setMidLineWidth(30)  # 设置线条中间部分的宽度
        #v_line.setStyleSheet("margin-top: 50px; margin-bottom: 50px;")  # 设置左右边距


        # 第二行第三列布局 - 脚本组件2
        work_widget = QWidget()
        work_layout = QVBoxLayout()
        self.Working_label = QLabel('未运行', self) # 制作 程序运行 显示表
        self.execute_script_button = QPushButton('执行脚本', self)
        self.stop_script_button = QPushButton('停止脚本', self)
        self.restart_script_button = QPushButton('重启脚本', self)

        work_layout.addWidget(self.Working_label)  # 添加显示列表到第二列布局
        work_layout.addWidget(self.execute_script_button)
        work_layout.addWidget(self.stop_script_button)
        work_layout.addWidget(self.restart_script_button)

        work_widget.setLayout(work_layout) #将 脚本组件2的布局 添加到 脚本组件2窗口

        #将调整好的各个窗口，按比例添加到 第二行布局中
        splitter.addWidget(test_widget)
        splitter.addWidget(v_line)
        splitter.addWidget(script_widget)
        splitter.addWidget(work_widget )

        splitter.setSizes([400, 10, 600, 400])  # 设置初始宽度比例
        second_row_layout.addWidget(splitter)

        # ========================================

        # 创建一个水平分割线
        h2_line = QFrame()
        h2_line.setFrameShape(QFrame.HLine)
        # h_line.setFrameShadow(QFrame.Sunken)  # 可选，设置阴影效果
        h2_line.setLineWidth(3)  # 设置线条宽度
        # h_line.setMidLineWidth(30)  # 设置线条中间部分的宽度
        h2_line.setStyleSheet("margin-top: 50px; margin-bottom: 50px;")  # 设置上下边距

        # ========================================

        # 第三行布局 - 待扩充
        third_row_layout = QHBoxLayout()
        # self.execute_script_button = QPushButton('执行脚本', self)
        # third_row_layout.addWidget(self.execute_script_button)

        # ========================================

        # 添加到主布局
        main_layout.addLayout(first_row_layout)
        main_layout.addWidget(h_line)   # 分割线
        main_layout.addLayout(second_row_layout)
        main_layout.addWidget(h2_line)  # 分割线
        main_layout.addLayout(third_row_layout)

        self.setLayout(main_layout) #显示整体布局

        # 恢复上次保存的 Excel 文件记录
        self.restore_excel_files()

        # ========================================
        # ========================================

        # 连接信号与槽
        self.leidian_folder_button.clicked.connect(self.select_leidian_folder)
        self.start_program_button.clicked.connect(self.start_program_and_app)

        self.load_test_image_button.clicked.connect(self.load_test_image)
        self.load_template_folder_button.clicked.connect(self.load_template_folder)
        self.match_image_button.clicked.connect(self.match_image)

        self.load_excel_button.clicked.connect(self.load_excel_file)
        self.delete_excel_button.clicked.connect(self.delete_selected_excel)
        self.move_up_button.clicked.connect(self.move_excel_up)
        self.move_down_button.clicked.connect(self.move_excel_down)

        self.execute_script_button.clicked.connect(self.execute_script)
        self.stop_script_button.clicked.connect(self.stop_script)
        self.restart_script_button.clicked.connect(self.restart_script)

# =================启动组件=======================================
    def select_leidian_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择雷电安装文件夹")
        if folder:
            self.leidian_path_label.setText(folder)
            self.selected_folder = folder
            self.load_apps_from_folder(folder)

    def load_apps_from_folder(self, folder_path):
        self.app_combobox.clear()  # 清空当前组合框
        apps_folder_path = os.path.join(folder_path, 'apps')

        # 添加“仅打开模拟器”的选项
        self.app_combobox.addItem("仅打开模拟器")

        if os.path.exists(apps_folder_path):
            for file_name in os.listdir(apps_folder_path):
                if file_name.endswith('.config'):
                    # 只显示文件名去掉 .config 扩展名的部分
                    display_name = os.path.splitext(file_name)[0]
                    self.app_combobox.addItem(display_name)

    def update_selected_app(self, index):
        self.selected_app = self.app_combobox.itemText(index)

    def start_program_and_app(self):
        # 检查 selected_folder 和 selected_app 是否已正确设置
        if not self.selected_folder:
            self.leidian_path_label.setText("请选择雷电安装文件夹")
            return  # 如果没有正确设置，提前返回，不再执行后续代码

        # 实例化控制台类
        self.dnconsole = Dnconsole(self.selected_folder)  # 使用选择的文件夹路径
        # 获取当前选中的App索引
        current_index = self.app_combobox.currentIndex()
        self.update_selected_app(current_index)
        if self.selected_app == "仅打开模拟器":
            # 如果选择了“仅打开模拟器”
            self.dnconsole.launch()  # 直接启动模拟器
        else:
            self.dnconsole.launchx(0, f'{self.selected_app}')

    def get_binding_handle(self, index=0):
        """
        从list2方法返回的结果中提取指定索引模拟器的绑定句柄
        :param result: list2方法返回的字符串
        :param index: 模拟器的索引，默认为0
        :return: 绑定句柄
        """
        result = self.dnconsole.list2()
        # 将结果按行分割
        lines = result.strip().split('\n')
        print(lines)
        if index >= len(lines):
            return None  # 如果索引超出范围，则返回None
        # 取出 指定索引的行(取出即指定模拟器的信息)，并按逗号分割
        parts = lines[index].split(',')
        # 返回绑定句柄，即第三个元素（索引为3）
        self.top_level_handle = parts[2]
        self.binding_handle = parts[3]
        print(f"绑定句柄: {self.binding_handle}")
        return self.binding_handle

    # def get_binding_handle(self):
    #     # 获取雷电模拟器子窗口的句柄
    #     if self.top_level_handle is None:
    #         self.top_level_handle = win32gui.FindWindow(None, "雷电模拟器")
    #     if self.top_level_handle:
    #         self.binding_handle = win32gui.FindWindowEx(self.top_level_handle, None, None, "子窗口标题")
    #     return self.binding_handle

    # ===============测试组件=====================
    def load_test_image(self):
        # 使用 folder_for_test_image 作为初始目录
        file, _ = QFileDialog.getOpenFileName(self, "读取测试大底图", self.folder_for_test_image, "Image Files (*.png *.jpg *.bmp)")
        if file:
            file_name = os.path.basename(file)
            self.test_image_path_label.setText(file_name)
            self.test_image_full_path = file
            print(f"选择了测试图片: {file_name}")
            # 更新 last_opened_folder
            self.folder_for_test_image = os.path.dirname(file)

    def load_template_folder(self):
        if self.dnconsole is None:
            text = "启动模拟器后，可以获取模拟器截图默认路径"
            self.template_folder_label.setText(text)
            print(text)
            folder = QFileDialog.getExistingDirectory(self, "读取模板图文件夹", self.folder_for_test_template)
        else:
            # 获取模拟器截图保存路径
            screenshot_path = self.dnconsole.images_path
            Sct_filename = os.path.basename(self.dnconsole.devicess_path)  # 获取模拟器截图图片文件名（固定的地址）
            template_folder_path = os.path.join(screenshot_path, Sct_filename)
            folder = QFileDialog.getExistingDirectory(self, "读取模板图文件夹", template_folder_path)

        if folder:
            self.template_folder_label.setText(folder)
            self.template_folder = folder
            print(f"选择了测试模板文件夹: {folder}")
            # 更新 last_opened_folder
            self.folder_for_test_template = folder

    def match_image(self):
        test_match.test_match_pic(self.test_image_full_path, self.template_folder)

    # ==============脚本Excell组件=====================
    def restore_excel_files(self):
        # 从 QSettings 中恢复 Excel 文件记录
        excel_files = self.settings.value("excel_files", [])
        for file in excel_files:
            item = QListWidgetItem(file)
            self.excel_list.addItem(item)

    def load_excel_file(self):
        # 这里应该是打开一个文件对话框让用户选择 Excel 文件
        # 并将选中的文件路径添加到 QListWidget
        # 示例代码中省略了实际的文件对话框
        file, _ = QFileDialog.getOpenFileName(self, "载入Excell文件", "", "Excel Files (*.xlsx *.xls)")
        item = QListWidgetItem(file)
        self.excel_list.addItem(item)
        self.save_excel_files()

    def delete_selected_excel(self):
        # 删除选中的 Excel 文件项
        for item in self.excel_list.selectedItems():
            self.excel_list.takeItem(self.excel_list.row(item))
        self.save_excel_files()

    def move_excel_up(self):
        # 将选中的 Excel 文件项向上移动
        selected_items = self.excel_list.selectedItems()
        for item in selected_items:
            row = self.excel_list.row(item)
            if row > 0:
                self.excel_list.takeItem(row)
                self.excel_list.insertItem(row - 1, item)
                self.excel_list.setCurrentItem(item)
        self.save_excel_files()

    def move_excel_down(self):
        # 将选中的 Excel 文件项向下移动
        selected_items = self.excel_list.selectedItems()
        for item in reversed(selected_items):
            row = self.excel_list.row(item)
            if row < self.excel_list.count() - 1:
                self.excel_list.takeItem(row)
                self.excel_list.insertItem(row + 1, item)
                self.excel_list.setCurrentItem(item)
        self.save_excel_files()

    def save_excel_files(self):
        # 保存当前 QListWidget 中的 Excel 文件记录到 QSettings
        excel_files = [self.excel_list.item(i).text() for i in range(self.excel_list.count())]
        self.settings.setValue("excel_files", excel_files)

    def closeEvent(self, event):
        # 当窗口关闭时，保存 QListWidget 中的 Excel 文件记录
        self.save_excel_files()
        super().closeEvent(event)
    # ==============脚本执行组件=====================
    def update_working_label(self, text):
        self.Working_label.setText(text)

    def on_script_finished(self):
        # 所有文件执行完毕后更新标签
        self.Working_label.setText("执行完成")
        # 更新 UI 的代码
        print("UI 已更新")

    def execute_script(self):
        # 创建并启动线程
        self.script_executor.dnconsole = self.dnconsole
        self.script_executor.selected_folder = self.selected_folder
        self.script_executor.excel_list = self.excel_list
        self.script_executor.Working_label = self.Working_label
        self.script_executor.binding_handle = self.binding_handle
        thread = threading.Thread(target=self.script_executor.execute_script_task)
        thread.start()

    def stop_script(self):
        self.script_executor.set_running(False)
        self.update_working_label("停止脚本")

    def restart_script(self):
        self.stop_script()
        self.script_executor.set_running(True)
        self.execute_script()
        self.update_working_label("重启脚本")

    def update_ui(self):
        # 定期检查任务状态并更新 UI
        if not self.running:
            self.update_working_label("停止脚本")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    ex.show()
    sys.exit(app.exec_())