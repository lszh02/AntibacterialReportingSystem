import os

import pyautogui
import time

current_path = os.path.dirname(__file__)


def mouse_click(img, click_times=1, l_or_r="left"):
    # 定义鼠标事件
    # pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159
    while True:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=click_times, interval=0.2, duration=0.2, button=l_or_r)
            break
        print("未找到匹配图片,0.5秒后重试")
        time.sleep(0.5)


def delete_record():
    num = 1
    while True:
        mouse_click(rf'{current_path}/delete.png')
        mouse_click(rf'{current_path}/enter.png')
        mouse_click(rf'{current_path}/enter.png')
        num += 1
        print(f'删除{num}条记录！')


if __name__ == '__main__':
    delete_record()
