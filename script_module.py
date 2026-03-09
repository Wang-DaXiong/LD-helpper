import time
import os
from pathlib import Path
import re

from operator import truediv

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
            # 模拟器截图
            dnconsole.screen_shot()
            image = cv2.imread(Sct_full_path)
            image_gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            # 进行匹配
            targeted_XY = Script_action.find_templates_in_image(image_gray, pic_templated_folder, 0.95)
            # dnconsole.actionOfTap(0, targeted_XY[0], targeted_XY[1])
            if len(targeted_XY) == 0:
                all_tapped = False  # 如果所有模板图没有找到，则设置标志为False
                break

            for (top_left, (h, w)) in targeted_XY:
                dnconsole.actionOfTap(0, top_left[0] + w / 2, top_left[1] + h / 2)
                time.sleep(3)  # 暂停2秒，等待点击动作完成

        return all_tapped  # 返回是否所有模板都找到了并点击了

    @staticmethod
    def find_and_tap_plus(dnconsole, Sct_full_path, pic_templated_folder, max_match_attempts=2):
        # 加入无法匹配时重复匹配机制，最多尝试 2次匹配图片
        m_attempt_count = 0
        match_success = False
        while m_attempt_count < max_match_attempts:
            m_attempt_count += 1
            # logger.info(f"尝试匹配图片，这是第{m_attempt_count}次尝试")
            print(f"正在匹配图片，这是第{m_attempt_count}次尝试")
            if Script_action.find_and_tap(dnconsole, Sct_full_path, pic_templated_folder):
                # logger.info("图片匹配成功")
                print("图片匹配成功")
                match_success = True
                break
            else:
                # logger.warning(f"存在图片匹配失败，已尝试{m_attempt_count}次")
                print(f"有图片匹配失败，已尝试{m_attempt_count}次，达到{max_match_attempts}次后将结束本轮匹配")
                time.sleep(2)  # 等待一段时间再重试
        # Script_action.recovery_operation(dnconsole)
        return match_success



#-------------------------------------------------------------------------
    @staticmethod
    def detect_image_and_escape(dnconsole, hwnd, Sct_full_path, pic_templated_folder, app_instance, interval=5):
        """
        监控屏幕上的指定图像，如果找到全部图像，则停止；如果没有找到全部图像，则每interval秒后重新检测。
        参数:
            hwnd (int): 窗口句柄。
            pic_templated_folder (str): 模板图像文件夹路径。
            interval (int): 每次检查之间的间隔时间（秒）。
        """
        template_paths = [os.path.join(pic_templated_folder, file_name) for file_name in
                          os.listdir(pic_templated_folder)]
        while True:
            if not app_instance.running:  # 检测是否中止执行脚本
                break
            # 模拟器截图 -> 获取大底图（灰度）
            dnconsole.screen_shot()
            BG_image = cv2.imread(Sct_full_path)
            background_image_gray = cv2.cvtColor(BG_image, cv2.COLOR_BGR2GRAY)  # 转换为灰度图像
            # 假设所有模板都匹配成功
            all_match = True
            # 创建线程池
            with ThreadPoolExecutor(max_workers=len(template_paths)) as executor:
                futures = {executor.submit(Script_action.match_image_for_detection, background_image_gray,
                                           templated_path): templated_path for templated_path in template_paths}
                for future in as_completed(futures):
                    templated_path = futures[future]
                    try:
                        result = future.result()
                        if result is None:
                            all_match = False
                            print(f"图像 {templated_path} 未找到。")
                    except Exception as e:
                        print(f"处理图像 {templated_path} 时发生错误: {e}")
                        all_match = False

            # 检查是否所有模板都被找到
            if all_match:
                print("所有图像均已找到，结束监控。")
                return
            else:
                # 如果没有找到所有图像，执行返回操作
                print("未找到所有监测图像，进行返回操作。")
                # win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_ESCAPE, 0)  # 发送按下ESC的WM_KEYDOWN消息
                # time.sleep(0.1)  # 短暂等待确保消息被处理
                # win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_ESCAPE, 0)  # 发送释放ESC的WM_KEYUP消息
                # time.sleep(interval)
                Dnconsole.actionOfKeyCode(dnconsole, 0, 111)
                time.sleep(interval)


    @staticmethod
    def detect_image_and_swipe(dnconsole, Sct_full_path, pic_templated_folder, app_instance,
                               interval=5, max_swipes_per_dir=5, swipe_counter=[0], direction=['down']):
        """
        在模拟器中循环检测图像，未找到时交替执行：
            - 'down': 中心 → 左上滑（原逻辑）
            - 'up':   中心 → 右下滑（真正的反向）
        """
        while app_instance.running:
            if not app_instance.running:
                print("脚本被手动终止")
                break

            # 截图
            dnconsole.screen_shot()
            BG_image = cv2.imread(Sct_full_path)
            if BG_image is None:
                print("截图加载失败")
                time.sleep(interval)
                continue

            background_image_gray = cv2.cvtColor(BG_image, cv2.COLOR_BGR2GRAY)

            # 获取模板图
            file_names = [f for f in os.listdir(pic_templated_folder)
                          if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            if not file_names:
                print("模板图文件夹为空")
                time.sleep(interval)
                continue

            pic_templated_list = [os.path.join(pic_templated_folder, f) for f in file_names]

            # 多线程匹配
            found = False
            with ThreadPoolExecutor(max_workers=4) as executor:
                futures = {
                    executor.submit(Script_action.match_image_for_detection, background_image_gray, tp): tp
                    for tp in pic_templated_list
                }
                for future in as_completed(futures):
                    try:
                        result = future.result()
                        if result:
                            templated_path = futures[future]
                            print(f"✅ 找到图像: {os.path.basename(templated_path)}")
                            for (top_left, (h, w)) in result:
                                center_x = top_left[0] + w / 2
                                center_y = top_left[1] + h / 2
                                dnconsole.actionOfTap(0, center_x, center_y)
                                time.sleep(1)
                            found = True
                            break
                    except Exception as e:
                        print(f"匹配错误: {e}")

            if found:
                # 重置状态，准备下一次搜索
                swipe_counter[0] = 0
                direction[0] = 'down'  # 可选：重置方向
                return

            # --- 执行滑动：交替方向 ---
            cx, cy = 640, 360  # 屏幕中心

            if direction[0] == 'down':
                # 原始方向：中心 → 左上
                to_x, to_y = 490, 210
                print(f"↖️  第 {swipe_counter[0] + 1} 次向左上滑动...")
                dnconsole.actionOfSwipe(cx, cy, to_x, to_y)
                swipe_counter[0] += 1

                if swipe_counter[0] >= max_swipes_per_dir:
                    direction[0] = 'up'
                    swipe_counter[0] = 0
                    print("🔄 切换为向右下滑动")

            elif direction[0] == 'up':
                # 反向：中心 → 右下（真正对称）
                to_x, to_y = 790, 510
                print(f"↘️  第 {swipe_counter[0] + 1} 次向右下滑动...")
                dnconsole.actionOfSwipe(cx, cy, to_x, to_y)
                swipe_counter[0] += 1

                if swipe_counter[0] >= max_swipes_per_dir:
                    direction[0] = 'down'
                    swipe_counter[0] = 0
                    print("🔄 切换为向左上滑动")

            time.sleep(interval)



    @staticmethod
    def detect_image_and_click(dnconsole, Sct_full_path, pic_templated_folder, app_instance, interval=5):
        while True:
            if not app_instance.running:  # 检测是否中止执行脚本
                break
            # 模拟器截图 -> 获取大底图（灰度）
            dnconsole.screen_shot()
            BG_image = cv2.imread(Sct_full_path)
            background_image_gray = cv2.cvtColor(BG_image, cv2.COLOR_BGR2GRAY)  # 转换为灰度图像

            # 构造每个图片的完整路径 -> 获取模板图（灰度）
            file_names = os.listdir(pic_templated_folder)
            pic_templated_list = [os.path.join(pic_templated_folder, file_name) for file_name in file_names]

            # 创建线程池
            with ThreadPoolExecutor(max_workers=len(pic_templated_list)) as executor:
                futures = {executor.submit(Script_action.match_image_for_detection, background_image_gray, templated_path): templated_path for templated_path in pic_templated_list}

                for future in as_completed(futures):
                    templated_path = futures[future]
                    try:
                        result = future.result()
                        if result is not None:
                            print(f"图像 {templated_path} 已找到，执行点击操作。")
                            for (top_left, (h, w)) in result:
                                dnconsole.actionOfTap(0, top_left[0] + w / 2, top_left[1] + h / 2)
                                time.sleep(2)  # 暂停2秒，等待点击动作完成
                            return
                    except Exception as e:
                        print(f"处理图像 {templated_path} 时发生错误: {e}")

            # 如果没有找到任何图像，等待下一次检测
            time.sleep(interval)


    @staticmethod
    def match_image_for_detection(background_image_gray, templated_path, threshold=0.95):
        """
        使用模板匹配检测图像是否存在于背景图像中。
        参数:
            BG_image_gray: 背景图像（灰度）(numpy.ndarray)
            templated_path: 模板图像文件路径 (str)
            threshold: 匹配度阈值 (float)
        返回:
            如果找到匹配的图像，返回的是匹配位置 max_loc 和模板图像的尺寸 (w, h)。
        """
        try:
            templated_image = cv2.imread(templated_path)
            if templated_image is None:
                print(f"未能加载图像: {templated_path}")
                return False
            templated_image_gray = cv2.cvtColor(templated_image, cv2.COLOR_BGR2GRAY)

            res = cv2.matchTemplate(background_image_gray, templated_image_gray, cv2.TM_CCOEFF_NORMED)

            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
            if max_val > threshold:
                h, w = templated_image_gray.shape[:2]  # 只获取高度和宽度
                return [(max_loc, (h, w))]  # 将结果封装成列表，以便于后续处理
            return None

        except Exception as e:
            print(f"发生错误: {e}")
            return False


# --------------------------------------------------------------------

    @staticmethod
    def execute_script_from_excel(dnconsole, hwnd_handle, excel_file_path, app_instance):
        # 功能：通过读取Excel内容来执行脚本
        # 获取模拟器截图保存路径
        screenshot_path = dnconsole.images_path
        # 获取模拟器截图-图片文件名（固定的地址）
        Sct_filename = os.path.basename(dnconsole.devicess_path)
        Sct_full_path = os.path.join(screenshot_path, Sct_filename)

        # 读取Excel文件
        df = pd.read_excel(excel_file_path)
        df_selected = df[['Action', '参数']]
        # 将DataFrame转换为字典列表
        # script = df.to_dict(orient='records')
        script = df_selected.to_dict(orient='records')

        # 当前工作目录
        # current_dir = Path(__file__).parent #获取当前ese执行文件所在的目录。
        current_dir = Path(os.getcwd()) #获取当前工作目录
        print(f"当前工作目录: {current_dir}")

        for step in script:
            if not app_instance.running:  # 检测是否中止执行脚本
                break
            action = step['Action']
            param = step['参数']

            if action == 'press_key':
                Script_action.send_key_to_LDwindow(hwnd_handle, param)
            elif action == 'wait':
                time.sleep(param)
            elif action == 'find_and_tap_image':
                param = param.replace('\\', '/')
                # 处理相对路径
                if not Path(param).is_absolute():
                    param = (current_dir / param).resolve().as_posix()
                Script_action.find_and_tap_plus(dnconsole, Sct_full_path, param)
            elif action == 'swipe':
                param = re.sub(r'[，,]', ',', param)
                # 解析坐标参数
                x0, y0, x1, y1 = map(int, param.split(','))
                dnconsole.actionOfSwipe(x0, y0, x1, y1)
                # dnconsole.actionOfSwipe(700, 300, 500, 300)  # 左右滑动
            elif action == 'tap':
                param = re.sub(r'[，,]', ',', param)
                x, y = map(int, param.split(','))
                dnconsole.actionOfTap(0, x, y)

            elif action == 'detect_image_and_Esc':
                param = param.replace('\\', '/')
                # 统一将所有类型的逗号转换为英文逗号
                param = re.sub(r'[，,]', ',', param)
                # 检查参数中是否包含时间间隔
                if ',' in param:
                    path, interval_str = param.split(',')
                    interval = int(interval_str.strip())
                else:
                    path = param
                    interval = 5  # 默认间隔时间，单位为秒
                # 处理相对路径
                if not Path(path).is_absolute():
                    path = (current_dir / path).resolve().as_posix()
                Script_action.detect_image_and_escape(dnconsole, hwnd_handle, Sct_full_path, path, app_instance, interval=interval)
            elif action == 'detect_image_and_swipe':
                param = param.replace('\\', '/')
                # 统一将所有类型的逗号转换为英文逗号
                param = re.sub(r'[，,]', ',', param)
                # 检查参数中是否包含时间间隔
                if ',' in param:
                    path, interval_str = param.split(',')
                    interval = int(interval_str.strip())
                else:
                    path = param
                    interval = 2  # 默认间隔时间，单位为秒
                # 处理相对路径
                if not Path(path).is_absolute():
                    path = (current_dir / path).resolve().as_posix()
                Script_action.detect_image_and_swipe(dnconsole, Sct_full_path, path, app_instance,
                                                        interval=interval)
            elif action == 'detect_image_and_click':
                param = param.replace('\\', '/')
                # 统一将所有类型的逗号转换为英文逗号
                param = re.sub(r'[，,]', ',', param)
                # 检查参数中是否包含时间间隔
                if ',' in param:
                    path, interval_str = param.split(',')
                    interval = int(interval_str.strip())
                else:
                    path = param
                    interval = 5  # 默认间隔时间，单位为秒
                # 处理相对路径
                if not Path(path).is_absolute():
                    path = (current_dir / path).resolve().as_posix()
                Script_action.detect_image_and_click(dnconsole, Sct_full_path, path, app_instance,
                                                     interval=interval)
            elif action == 'press_key_Esc':
                #顶层句柄，用模拟器 Esc强制返回
                Dnconsole.actionOfKeyCode(dnconsole, 0, 111)




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



