import os
import cv2

from LD_dnconsole import Dnconsole
from script_module import Script_action

class test_match:
    def __init__(self):
        pass
    @staticmethod
    def test_match_pic(Background_pic_full_path, pic_templated_folder):
        # 读取图片的完整路径， 转为opencv的数据内容
        image = cv2.imread(Background_pic_full_path)
        image_gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # 转为灰度图，提高匹配速度

        targeted_XY = Script_action.find_templates_in_image(image_gray, pic_templated_folder, 0.95)

        # 根据匹配的坐标，绘制矩形
        for (top_left, (h, w)) in targeted_XY:
            print(f'{top_left},{(h, w)}')
            bottom_right = (top_left[0] + w, top_left[1] + h)
            cv2.rectangle(image, top_left, bottom_right, (0, 255, 0), 2)  # 绿色框，线宽2

        # 显示缩小后的图片
        cv2.imshow('Test_Match_Pic', image)
        # 按任意键关闭显示窗口
        cv2.waitKey(0)
        cv2.destroyAllWindows()

    @staticmethod
    def test_match_tap(dnconsole, pic_templated_folder):
        # 获取模拟器截图保存路径
        screenshot_path = dnconsole.images_path
        # 获取模拟器截图图片文件名（固定的地址）
        Sct_filename = os.path.basename(dnconsole.devicess_path)
        Sct_full_path = os.path.join(screenshot_path, Sct_filename)
        # 读取图片的完整路径， 转为opencv的数据内容
        image = cv2.imread(Sct_full_path)
        image_gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # 转为灰度图，提高匹配速度

        targeted_XY = Script_action.find_templates_in_image(image_gray, pic_templated_folder, 0.95)
        bool = Script_action.find_and_tap_plus(dnconsole, Sct_full_path, pic_templated_folder)
        return  bool


    @staticmethod
    def get_binding_handle_for_test(result, index=0):
        """
        从list2方法返回的结果中提取指定索引模拟器的绑定句柄
        :param result: list2方法返回的字符串
        :param index: 模拟器的索引，默认为0
        :return: 绑定句柄
        """
        # 将结果按行分割
        lines = result.strip().split('\n')
        print(lines)

        if index >= len(lines):
            return None  # 如果索引超出范围，则返回None

        # 取出指定索引的行，并按逗号分割
        parts = lines[index].split(',')

        # 返回绑定句柄，即第三个元素（索引为3）
        return parts[2], parts[3]



if __name__ == '__main__':
    # ============识图并点击测试=================
    # 实例化控制台类
    dnconsole = Dnconsole(r'D:\leidian\LDPlayer9')  # 替换为实际安装路径
    # 启动模拟器
    #dnconsole.launch()

    # 获取模拟器截图图片文件名（固定的地址）
    Sct_full_path = r'C:\Users\76703\Documents\leidian9\Pictures\Screenshots\screenshot_tmp.png'

    # 获取模板图片文件夹（输入固定的地址）
    # 模板图片要有特征、小一点
    pic_templated_folder = r'D:\Programming Study\0908-Fist-try\PIC-From LD\test'

    # 读取图片的完整路径， 转为opencv的数据内容
    image = cv2.imread(Sct_full_path)
    image_gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # 转为灰度图，提高匹配速度

    targeted_XY = Script_action.find_templates_in_image(image_gray, pic_templated_folder, 0.95)
    bool = Script_action.find_and_tap_plus(dnconsole,Sct_full_path,pic_templated_folder)
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

