import os
import logging
import subprocess

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class Dnconsole:
    """
    【雷电控制台类】
    version: 9.0
    import该文件会自动实例化为 Dc
    """

    def __init__(self, installation_path: str):
        """
        【构造方法】
        """
        # if 模拟器安装路径存在性检测
        if not os.path.exists(installation_path):
            raise FileNotFoundError(f'模拟器安装路径不存在！{installation_path}')
        # 获取模拟器安装路径
        self.ins_path = installation_path
        # -------------------------------------------------------
        # Dnconsole程序路径
        self.console_path = self.ins_path + r'\ldconsole.exe '
        if not os.path.exists(self.console_path):
            raise FileNotFoundError(f'程序路径不存在！请确认模拟器安装文件是否完整！{self.console_path}')
        # adb程序路径
        self.adb_path = self.ins_path + r'\adb.exe '
        if not os.path.exists(self.adb_path):
            raise FileNotFoundError(f'程序路径不存在！请确认模拟器安装文件是否完整！{self.adb_path}')
        # ld程序路径
        self.ld_path = self.ins_path + r'\ld.exe '
        # -------------------------------------------------------
        # 模拟器截屏程序路径
        self.screencap_path = r'/system/bin/screencap'
        # 模拟器截图保存路径
        self.devicess_path = r'/sdcard/Pictures/Screenshots/screenshot_tmp.png'
        # 用户环境
        self.userprofile = os.environ['USERPROFILE']
        self.documents_path = os.path.join(self.userprofile, 'Documents')
        # 本地图片保存路径
        self.images_path = self.documents_path + r'\leidian9\Pictures\Screenshots\\'

        # if 模拟器的截图保存路径
        if os.path.exists(self.images_path) is False:
            print('模拟器截图保存路径不存在！')

        # # 本地 '样本图片' 的保存路径
        # self.workspace_path = os.getcwd()   # 获取当前工作目录路径
        # self.target_path = self.workspace_path + r'\PIC-From LD\\'

        logger.info('Class-Dnconsole is ready. (%s)', self.ins_path)

#==============================================
    def CMD(self, cmd: str):
        """
        【执行控制台命令语句】
        :param cmd: 命令
        :return: 控制台调试内容
        """
        CMD = self.console_path + cmd  # 控制台命令
        process = os.popen(CMD)
        result = process.read()
        print(result)
        process.close()
        return result

    def ldCMD(self, cmd: str):
        """
        【通过ld.exe执行代替adb的命令语句】
        :param cmd: 命令
        :return: 控制台调试内容
        """
        ldCMD = self.ld_path + cmd  # 控制台命令
        # print(f"Executing command: {ldCMD}")  # 调试输出
        process = os.popen(ldCMD)
        result = process.read()
        process.close()
        # print(f"Command result: {result}")  # 调试输出
        return result

    def ADB(self, cmd: str):
        """
        【执行ADB命令语句】
        :param cmd: 命令
        :return: 控制台调试内容
        """
        CMD = self.adb_path + cmd  # adb命令
        # p = subprocess.Popen(CMD, stdout=subprocess.PIPE)
        p = subprocess.Popen(CMD, stdout=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW)
        output, _ = p.communicate()
        return output

#======模拟器=======================

    def launch(self, index: int = 0):
        '''
        启动模拟器
        :param index: 模拟器序号
        :return: True=已启动 / False=不存在
        '''
        cmd = f'launch --index {index}'
        logger.info(f'Executing command: {cmd}')

        result = self.CMD(cmd)

        if result.strip() == '':
            logger.info(f'Simulator launched successfully with index {index}.')
            return True
        else:
            logger.error(f'Failed to launch simulator with index {index}. Output: {result}')
            return False

    def isrunning(self, index: int = 0):
        """
        【检测模拟器是否启动】
        :param index: 模拟器序号
        :return: True=已启动 / False=未启动
        """
        cmd = 'isrunning --index %d' % index
        if self.CMD(cmd) == 'running':
            return True
        else:
            return False

    def launchx( self, index:int, packagename:str ):
        '''
        【同时启动模拟器和App】
        :param index: 模拟器序号
        :param packagename: 包名
        :return: 控制台调试内容
        '''
        cmd = 'launchex --index %d --packagename "%s"' %(index, packagename)
        return self.CMD(cmd)


 #========模拟器之APP=======================

    def runApp(self, index: int, packagename: str):
        """
        【运行App】
        :param index: 模拟器序号
        :param packagename: 包名
        :return: 控制台调试内容
        """
        cmd = 'runapp --index %d --packagename %s' % (index, packagename)
        return self.CMD(cmd)

# =========模拟器截图========================

    def screen_shot(self):
        #cmd = f'{self.adb_path} -s emulator-5554 shell {self.screencap_path} -p {self.devicess_path}'
        cmd = f' -s emulator-5554 shell {self.screencap_path} -p {self.devicess_path}'
        self.ADB(cmd)
        return None
    
    def screenget(self, name: str):  # 截屏,并存放到指定路径
        local_path = f'{self.images_path}{name}.png'
        cmd = f'{self.adb_path} -s emulator-5554 pull {self.devicess_path} {local_path}'
        process = os.popen(cmd)
        result = process.read()
        process.close()
        #print(result)
        return result


# =========模拟器键鼠操作========================
    def list2(self):    #获取句柄，为后台键盘按钮作准备
        '''
        【取模拟器列表】
        :return: 列表（索引、标题、顶层句柄、绑定句柄、是否进入android、进程PID、VBox进程PID）
        '''
        cmd = 'list2'
        return self.CMD(cmd)

    def actionOfTap( self, index:int, x:int, y:int ) -> str:
        '''
        【点击操作】
        :param index: 模拟器序号
        :param x: x
        :param y: y
        :return: 控制台调试内容
        '''
        cmd = 'adb --index %d --command "shell input tap %d %d"' %(index, x, y)
        return self.CMD(cmd)

    def actionOfTap_ADB( self, x:int, y:int ) -> str:
        '''
        【点击操作】 跟上面的def actionOfTap方法一样，只是写法不同而已
        :param index: 模拟器序号
        :param x: x
        :param y: y
        :return: 控制台调试内容
        '''
        cmd = f' -s emulator-5554 shell input tap {x} {y}'
        return self.ADB(cmd)

    # =======以下只对模拟器作用，模拟器的子窗口无效传达========

    def actionOfTap_Ld(self, index: int, x: int, y: int):
        """
        【点击操作】
        :param index: 模拟器序号
        :param x: x
        :param y: y
        :return: 控制台调试内容
        """
        # cmd = '-s %d input tap %d %d' % (index, x, y)
        # 使用 f-string 构建命令字符串
        cmd = f'-s {index} input tap {x} {y}'
        return self.ldCMD(cmd)

    def actionOfSwipe( self, x0:int, y0:int, x1:int, y1:int, index=0, ms:int = 200 ):
        '''
        【滑动操作】
        :param index: 模拟器序号
        :param x0,y0: 起点坐标
        :param x1,y1: 终点坐标
        :param ms: 滑动时长
        :return: 控制台调试内容
        '''
        cmd = 'adb --index %d --command "shell input swipe %d %d %d %d %d"' %(index, x0, y0, x1, y1, ms)
        return self.CMD(cmd)
        # dnconsole.actionOfSwipe(0, 700, 300, 500, 300) #左右滑动

    def actionOfKeyCode(self, index: int,keycode: str):
        '''
        【键码输入操作, 向指定的安卓模拟器发送按键事件】 #该方法只会发送给 顶层句柄, 子窗口无效
        :param index: 模拟器序号
        :param keycode: 键码（0~9，10=空格，111=ESC）
        :return: 控制台调试内容

        安卓键码对照表(只需要输入str内容即可)：'KEYCODE_space': 62，'KEYCODE_D': 32......
        '''
        try:
            cmd = 'adb --index %d --command "shell input keyevent %s"' % (index, keycode)
            logging.info(f"Executing command: {cmd}")  # 调试输出
            return self.CMD(cmd)
        except Exception as e:
            return str(e)

    # def action_of_keyboard(self, keyboard_value: str, index=0):
    #     """
    #     模拟键盘操作。
    #     :param keyboard_value: 按键值，可以是 'back', 'home', 'menu', 'volumeup', 'volumedown'
    #     :param index: 设备索引，默认为0
    #     :return: 执行命令的结果
    #     """
    #     # 构建命令字符串
    #     # cmd = 'action --index %d --key call.keyboard --value %s' % (index, key)
    #     cmd = f'action --index {index} --key call.keyboard --value {keyboard_value}'
    #
    #     try:
    #         # 调用CMD方法执行命令
    #         result = self.CMD(cmd)
    #     except Exception as e:
    #         # 处理可能出现的异常
    #         print(f"An error occurred while executing the command: {e}")
    #         result = None
    #
    #     return result

    def acionOfKeyboard_LD(self, Keydcode, index=0):
        #该方法只会发送给 顶层句柄, 子窗口无效
        cmd = f'-s {index} input keyevent {Keydcode}'
        return self.ldCMD(cmd)

    def actionOfInput(self, index: int, text: str):
        '''
        【输入操作】
        :param index: 模拟器序号
        :param text: 文本内容
        :return: 控制台调试内容
        '''
        cmd = 'action --index %d --key call.input --value "%s"' % (index, text)
        return self.CMD(cmd)

