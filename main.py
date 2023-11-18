from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from pptx import Presentation
from PIL import Image
import io
import subprocess


# 设置下载路径
directory = '/Users/sines/Downloads/ppts/Luck PPT'
output_path = './output'
compressed_output_path = './compressed'

# 设置webdriver路径
webdriver_path = './msedgedriver'


def convert_ppt_to_pptx(ppt_path):
    script = f'''
    tell application "Microsoft PowerPoint"
        open "{ppt_path}"
        set pptx_path to "{os.path.splitext(ppt_path)[0]}.pptx"
        save active presentation as save as presentation in pptx_path file format PPTX file format
        close active presentation
    end tell
    '''
    subprocess.run(["osascript", "-e", script])



def translate(file_path):
# 查找元素并执行操作
    try:
        # 创建webdriver实例
        edge_options = webdriver.EdgeOptions()
        prefs = {"download.default_directory": output_path}
        edge_options.add_experimental_option("prefs", prefs)
        driver = webdriver.Edge(options=edge_options)
        driver.get('https://translate.google.com/?hl=zh-CN&sourceid=cnhp&sl=en&tl=zh-TW&op=docs')
        # 找到上传文件的按钮并上传文件
        file_upload = driver.find_element(By.ID, 'ucj-19')
        file_upload.send_keys(file_path)        

        # 等待翻译按钮出现
        translate_button_xpath = '/html/body/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div[2]/div/div/button'
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, translate_button_xpath)))

        # 找到翻译按钮并点击
        submit_button = driver.find_element(By.XPATH, translate_button_xpath)
        submit_button.click()

        time.sleep(10)  # 等待5秒

        WebDriverWait(driver, 100).until(EC.text_to_be_present_in_element((By.XPATH, '/html/body/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div[2]/div/button/span[2]'), '下载译文'))

        # 点击下载按钮下载到 output_path
        download_button_xpath = '/html/body/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div[2]/div/button'
        download_button = driver.find_element(By.XPATH, download_button_xpath)
        download_button.click()
    finally:
        # 完成后关闭浏览器
        time.sleep(2)  # 等待5秒
        driver.quit()

# 设置文件大小限制（10MB）
size_limit = 10 * 1024 * 1024  # 10MB in bytes
for filename in os.listdir(directory):
    if filename.endswith(".ppt"):
        file_path = os.path.join(directory, filename)
        convert_ppt_to_pptx(file_path)

for filename in os.listdir(directory):
    if filename.endswith(".pptx"):
        file_path = os.path.join(directory, filename)
        file_size = os.path.getsize(file_path)
        if file_size > size_limit:
            print(f"文件 {filename} 的大小超过10MB")
            continue
        print('Begin translate:' + file_path)
        translate(file_path)