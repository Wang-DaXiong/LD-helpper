import time
import os
import numpy as np
import pandas as pd
import json

import win32api
import win32con

import cv2
import pyautogui

import threading
from queue import Queue  #Queue：用于在线程间安全地传递数据。每个线程将匹配结果放入队列中，主程序从队列中取出所有结果。
from concurrent.futures import ThreadPoolExecutor, as_completed #线程池来限制同时运行的线程数量

from LD_dnconsole import Dnconsole

class Script_action:
    def __init__(self):
        pass

    # @staticmethod
    # def find_templates_in_image(pic_background, pic_templated_folder, threshold=0.95 ):
    #     """
    #     采用 "单线程 "，在给定的大图中寻找多个模板图的位置。
    #     参数:
    #     pic_background (numpy.ndarray): 背景大图。
    #     pic_templated_folder (str): 模板图所在文件夹的路径。
    #     threshold (float): 匹配度阈值，默认为0.85。
    #     返回:
    #     list: 每个模板图像对应的匹配区域的坐标列表
    #     """
    #     # 构造每个图片的完整路径
    #     file_names = os.listdir(pic_templated_folder)  #读取文件里所有子文件名
    #     pic_templated_list = [os.path.join(pic_templated_folder, file_name) for file_name in file_names]
    #
    #     result_queue = Queue()
    #     threads = []
    #     for template_path in pic_templated_list:  # 依次读取模板图像
    #         pic_templated = cv2.imread(template_path)
    #         if pic_templated is None:
    #             print(f"未能加载图像: {template_path}")
    #             continue
    #         # 将模板图像转换为灰度模式
    #         pic_templated_gray = cv2.cvtColor(pic_templated, cv2.COLOR_BGR2GRAY)
    #
    #         thread = threading.Thread(target=Script_action.match_pic, args=(pic_background,pic_templated_gray,threshold, result_queue))
    #         thread.start()
    #         threads.append(thread)
    #
    #     for thread in threads:
    #         thread.join() #使用 thread.join() 确保所有线程完成后再继续执行主程序
    #
    #     # 从队列中收集所有结果
    #     match_XY_list = []
    #     while not result_queue.empty():
    #         match_XY_list.append(result_queue.get()) #使用 Queue，避免了多个线程同时写入同一个列表时的竞态条件问题。
    #
    #     return match_XY_list

    @staticmethod
    def find_templates_in_image(pic_background, pic_templated_folder, threshold=0.95, max_workers=4):
        """
        采用 "线程池 "，在给定的大图中寻找多个模板图的位置。
        参数:
        pic_background (numpy.ndarray): 背景大图。
        pic_templated_folder (str): 模板图所在文件夹的路径。
        threshold (float): 匹配度阈值，默认为0.95。
        max_workers (int): 线程池的最大线程数，默认为4。
        返回:
        list: 每个模板图像对应的匹配区域的坐标列表
        """
        # 构造每个图片的完整路径
        file_names = os.listdir(pic_templated_folder)
        pic_templated_list = [os.path.join(pic_templated_folder, file_name) for file_name in file_names]

        result_queue = Queue()

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = []
            for template_path in pic_templated_list:
                pic_templated = cv2.imread(template_path)
                if pic_templated is None:
                    print(f"未能加载图像: {template_path}")
                    break
                pic_templated_gray = cv2.cvtColor(pic_templated, cv2.COLOR_BGR2GRAY)
                futures.append(executor.submit(Script_action.match_pic, pic_background, pic_templated_gray, threshold,
                                               result_queue))

            # 等待所有任务完成
            for future in as_completed(futures):
                future.result()  # 这里可以捕获并处理任何异常

        # 从队列中收集所有结果
        match_XY_list = []
        while not result_queue.empty():
            match_XY_list.append(result_queue.get())

        return match_XY_list

    @staticmethod
    def match_pic(pic_background, pic_templated_gray, threshold, result_queue):
        # 使用模板匹配
        match_result = cv2.matchTemplate(pic_background, pic_templated_gray, cv2.TM_CCOEFF_NORMED)
        # 获取匹配度大于等于阈值的坐标
        match_result2 = np.where(match_result >= threshold)
        # 一张背景图可能有匹配多个模板图，从而返回多个坐标值
        match_result_XY = list(zip(*match_result2[::-1]))
        if len(match_result_XY) > 0:
            for pt in match_result_XY:
                result_queue.put((pt, pic_templated_gray.shape[:2]))  # 保存坐标及模板大小
    @staticmethod
    def find_and_tap(dnconsole, Sct_full_path, pic_templated_folder):
        all_tapped = True  # 假设所有模板都能找到并点击
        for i in range(len(os.listdir(pic_templated_folder))):
            dnconsole.screen_shot()
            image = cv2.imread(Sct_full_path)
            image_gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

            targeted_XY = Script_action.find_templates_in_image(image_gray, pic_templated_folder, 0.95)
            #dnconsole.actionOfTap(0, targeted_XY[0], targeted_XY[1])
            if len(targeted_XY) == 0:
                all_tapped = False  # 如果任何一个模板没有找到，则设置标志为False
                break

            for (top_left, (h, w)) in targeted_XY:
                dnconsole.actionOfTap(0, top_left[0] + w / 2, top_left[1] + h / 2)
                time.sleep(2)  # 暂停2秒，等待点击动作完成

        return all_tapped  # 返回是否所有模板都找到了并点击了
    @staticmethod
    def find_and_tap_plus(dnconsole, Sct_full_path, pic_templated_folder, max_match_attempts=2):
        # 加入无法匹配时重复匹配机制，最多尝试 2次匹配图片
        m_attempt_count = 0
        match_success = False
        while m_attempt_count < max_match_attempts:
            m_attempt_count += 1
            #logger.info(f"尝试匹配图片，这是第{m_attempt_count}次尝试")
            print(f"正在匹配图片，这是第{m_attempt_count}次尝试")
            if Script_action.find_and_tap(dnconsole, Sct_full_path, pic_templated_folder):
                #logger.info("图片匹配成功")
                print("图片匹配成功")
                match_success = True
                break
            else:
                #logger.warning(f"存在图片匹配失败，已尝试{m_attempt_count}次")
                print(f"有图片匹配失败，已尝试{m_attempt_count}次，达到{max_match_attempts}次后将结束本轮匹配")
                time.sleep(2)  # 等待一段时间再重试
        #Script_action.recovery_operation()
        return match_success

    @staticmethod
    def recovery_operation():
        print("返回初始界面，等待重新匹配图片")
        pyautogui.press('esc', presses=2, interval=0.5)  # 按三次ESC键
        # win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vkey, 0)  # 发送 WM_KEYDOWN 消息
        # time.sleep(0.1)  # 确保消息被处理
        # win32api.PostMessage(hwnd, win32con.WM_KEYUP, vkey, 0)  # 发送 WM_KEYUP 消息
        time.sleep(2)  # 等待恢复操作生效

    @staticmethod
    def execute_script_from_Py(dnconsole, script):
    #通过读取字典执行，字典需要有一定格式
        screenshot_path = dnconsole.images_path
        Sct_filename = os.path.basename(dnconsole.devicess_path)
        Sct_full_path = os.path.join(screenshot_path, Sct_filename)

        for step in script:
            action = step['action']
            if action == 'press_key':
                key = step['键盘按键']
                # delay = step.get('delay', 0)  # 使用get方法以避免键不存在时报错
                pyautogui.press(key)
                # time.sleep(delay)
            elif action == 'find_and_tap_image':
                template_folder = step['模板文件夹']
                Script_action.find_and_tap_plus(dnconsole, Sct_full_path, template_folder)
            elif action == 'wait':
                duration = step['时间(秒)']
                print(f'等待 {duration} 秒')
                time.sleep(duration)
            elif action == 'swip':
                x0, y0, x1, y1 = step['左右1'], step['上下1'], step['左右2'], step['上下2'],
                dnconsole.actionOfSwipe(0, x0, y0, x1, y1)
                # dnconsole.actionOfSwipe(0, 700, 300, 500, 300)  # 左右滑动
    @staticmethod
    def execute_script_from_txt(dnconsole, file_path):
        # 通过读取字txt文件执行脚本，需要有一定格式

        # 获取模拟器截图保存路径
        screenshot_path = dnconsole.images_path
        # 获取模拟器截图图片文件名（固定的地址）
        Sct_filename = os.path.basename(dnconsole.devicess_path)
        Sct_full_path = os.path.join(screenshot_path, Sct_filename)

        # 从文件中读取 JSON 数据
        with open(file_path, 'r', encoding='utf-8') as file:
            script_json = file.read()
        # 将 JSON 字符串转换为 Python 列表
        script_list = json.loads(script_json)

        # 执行脚本中的每个步骤
        for step in script_list:
            if step['step'] == 'press_key':
                pyautogui.press(step['键盘按键'])
            elif step['step'] == 'wait':
                time.sleep(step['时间(秒)'])
                print(f'等待 {step['时间(秒)']} 秒')
            elif step['step'] == 'find_and_tap_image':
                Script_action.find_and_tap_plus(dnconsole, Sct_full_path, step['模板文件夹'])

    @staticmethod
    def execute_script_from_excel_test(dnconsole, hwnd_handle, Excell_file_path):
        #通过读取Excell内容来执行脚本
        # 获取模拟器截图保存路径
        screenshot_path = dnconsole.images_path
        # 获取模拟器截图图片文件名（固定的地址）
        Sct_filename = os.path.basename(dnconsole.devicess_path)
        Sct_full_path = os.path.join(screenshot_path, Sct_filename)

        # 读取Excel文件
        df = pd.read_excel(Excell_file_path)
        # 将DataFrame转换为字典列表
        script = df.to_dict(orient='records')

        for step in script:
            action = step['Action']
            param = step['参数']

            if action == 'press_key':
                Script_action.send_key_to_LDwindow(hwnd_handle, param)
                print('按键中')
            elif action == 'wait':
                time.sleep(param)
                print('等待中')
            elif action == 'find_and_tap_image':
                Script_action.find_and_tap_plus(dnconsole, Sct_full_path, param)
                print('匹图中')
    @staticmethod
    def execute_script_from_excel(dnconsole, hwnd_handle, Excell_file_path, app_instance):
    #功能：通过读取Excell内容来执行脚本
        # 获取模拟器截图保存路径
        screenshot_path = dnconsole.images_path
        # 获取模拟器截图图片文件名（固定的地址）
        Sct_filename = os.path.basename(dnconsole.devicess_path)
        Sct_full_path = os.path.join(screenshot_path, Sct_filename)

        # 读取Excel文件
        df = pd.read_excel(Excell_file_path)
        # 将DataFrame转换为字典列表
        script = df.to_dict(orient='records')

        for step in script:
            if not app_instance.running: #检测是否中止执行脚本
                break
            action = step['Action']
            param = step['参数']

            if action == 'press_key':
                Script_action.send_key_to_LDwindow(hwnd_handle, param)
            elif action == 'wait':
                time.sleep(param)
            elif action == 'find_and_tap_image':
                param = param.replace('\\', '/')
                Script_action.find_and_tap_plus(dnconsole, Sct_full_path, param)
            elif action == 'swipe':
                # 解析坐标参数
                x0, y0, x1, y1 = map(int, param.split(','))
                dnconsole.actionOfSwipe(x0, y0, x1, y1)
                # dnconsole.actionOfSwipe(700, 300, 500, 300)  # 左右滑动

    @staticmethod
    def char_to_vkey(char):
        # 特殊字符映射
        special_keys = {
            ' ': win32con.VK_SPACE,
            '\t': win32con.VK_TAB,
            '\n': win32con.VK_RETURN,
            '\b': win32con.VK_BACK,
            '\r': win32con.VK_RETURN,
            '-': 0xBD,  # VK_OEM_MINUS
            '=': 0xBB,  # VK_OEM_PLUS
            '[': 0xDB,  # VK_OEM_4
            ']': 0xDD,  # VK_OEM_6
            '\\': 0xDC,  # VK_OEM_5
            ';': 0xBA,  # VK_OEM_1
            '\'': 0xDE,  # VK_OEM_7
            ',': 0xBC,  # VK_OEM_COMMA
            '.': 0xBE,  # VK_OEM_PERIOD
            '/': 0xBF,  # VK_OEM_2
            '`': 0xC0,  # VK_OEM_3
            'esc': win32con.VK_ESCAPE  # 添加对 ESC 键的支持
        }

        # 首先尝试从特殊字符映射中获取虚拟键码
        vkey = special_keys.get(char.lower(), None)

        # 如果没有找到，则检查是否为单个字符
        if vkey is None:
            if len(char) != 1:
                raise ValueError("只能转换单个字符")
            # 字母和数字可以直接用其 ASCII 值
            if char.isalnum():
                return ord(char.upper())  # 大写处理，因为虚拟键码区分大小写
            else:
                raise ValueError(f"无法识别的字符: {char}，请在ScriptAction.char_to_vkey方法中添加转译")
        return vkey

    @staticmethod
    def send_key_to_LDwindow(hwnd, key_char):
        try:
            # 把信息转 ASCII码值（十进制）
            vkey = Script_action.char_to_vkey(key_char)
            win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vkey, 0)  # 发送 WM_KEYDOWN 消息
            time.sleep(0.1)  # 确保消息被处理
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, vkey, 0)  # 发送 WM_KEYUP 消息
        except ValueError as e:
            print(f"错误: {e}")   # 可以在这里记录日志或者进行其他错误处理


if __name__ == '__main__':
    # 启动模拟器 和 App
    #dnconsole.launch()  # 启动模拟器
    # dnconsole.launchx(0, 'com.tencent.tmgp.maplem')
    # time.sleep(5)  # 等待模拟器启动

    # # 获取模拟器截图保存路径
    # screenshot_path = dnconsole.images_path
    # # 获取模拟器截图图片文件名（固定的地址）
    # Sct_filename = os.path.basename(dnconsole.devicess_path)
    # Sct_full_path = os.path.join(screenshot_path, Sct_filename)


    # ============识图并点击测试=================
    # 实例化控制台类
    dnconsole = Dnconsole(r'D:\leidian\LDPlayer9')  # 替换为实际安装路径
    # 启动模拟器
    dnconsole.launch()

    # 获取模拟器截图图片文件名（固定的地址）
    Sct_full_path = r'C:\Users\76703\Documents\leidian9\Pictures\Screenshots\screenshot_tmp.png'
    # 获取模板图片文件夹（输入固定的地址）
    # 模板图片要有特征、小一点、
    pic_templated_folder = r'D:\Programming Study\0908-Fist-try\PIC-From LD\test'

    # 读取图片的完整路径， 转为opencv的数据内容
    image = cv2.imread(Sct_full_path)
    image_gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # 转为灰度图，提高匹配速度

    targeted_XY = Script_action.find_templates_in_image(image_gray, pic_templated_folder, 0.95)
    bool = Script_action.find_and_tap(dnconsole,Sct_full_path,pic_templated_folder)
    print(bool)

    # =====可视化内容，用于测试=========
    # 根据匹配的坐标，绘制矩形
    for (top_left, (h, w)) in targeted_XY:
        print(f'{top_left},{(h, w)}')
        bottom_right = (top_left[0] + w, top_left[1] + h)
        cv2.rectangle(image, top_left, bottom_right, (0, 255, 0), 2)  # 绿色框，线宽2

    # 显示缩小后的图片
    cv2.imshow('Screenshot', image)
    cv2.waitKey(0)
    cv2.destroyAllWindows()



