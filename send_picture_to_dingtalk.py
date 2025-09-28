import os
import time
import pyautogui
import pyperclip

# 获取当前用户的主目录
pwd_dir = os.path.expanduser("/tmp/office_automation/")
# 图片目录路径（请根据实际情况调整）
PICTURE_DIR = os.path.join(pwd_dir, "picture")

# MacOS钉钉相关坐标位置（需要根据实际情况调整）
DINGTALK_POS = (330, 50)          # 钉钉位置
SEARCH_BOX_POS = (710, 43)        # 搜索框位置
ATTACH_BUTTON_POS = (709, 1395)   # 附件按钮位置
ATTACH_FILE_POS = (710, 1259)     # 附件按钮位置
SEND_BUTTON_POS = (1195, 1393)    # 发送按钮位置

def send_image_to_dingtalk(name, image_path):
    """发送图片给指定联系人"""
    try:
        # 激活钉钉
        pyautogui.click(DINGTALK_POS)
        time.sleep(2)  # 等待钉钉启动

        # 激活搜索框（Command+f）
        pyautogui.hotkey('command', 'f')
        time.sleep(2)

        # 输入联系人姓名
        pyperclip.copy(name)
        pyautogui.hotkey('command', 'v')
        time.sleep(2)

        # 回车选择第一个搜索结果
        pyautogui.press('enter')
        time.sleep(2)

        # 点击上传附件按钮
        pyautogui.click(ATTACH_BUTTON_POS)
        time.sleep(2)

        # 点击上传附件->上传文件
        pyautogui.click(ATTACH_FILE_POS)
        time.sleep(3)

        pyperclip.copy(image_path)
        # macOS打开路径输入框
        pyautogui.hotkey('command', 'shift', 'g')
        time.sleep(3)
        # 输入文件路径
        pyautogui.hotkey('command', 'v')
        # 选中文件
        pyautogui.press('enter')
        time.sleep(3)

        # 点击确定
        pyautogui.press('enter')
        time.sleep(2)

        # 发送图片
        pyautogui.click(SEND_BUTTON_POS)

        print(f"已发送图片给 {name}")
        return True

    except Exception as e:
        print(f"发送给 {name} 失败: {str(e)}")
        return False

def main():
    # 确保图片目录存在
    if not os.path.exists(PICTURE_DIR):
        print(f"图片目录不存在: {PICTURE_DIR}")
        return

    print("=== MacOS通过钉钉自动发送图片 ===")
    print("请在5秒内将钉钉窗口置于前台...")
    time.sleep(5)

    # 遍历图片目录
    for filename in os.listdir(PICTURE_DIR):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            # 从文件名提取姓名（去掉扩展名）
            name = os.path.splitext(filename)[0]
            image_path = os.path.join(PICTURE_DIR, filename)

            print(f"准备发送给: {name}")
            send_image_to_dingtalk(name, image_path)
            # 发送间隔
            time.sleep(10)

if __name__ == "__main__":
    main()
