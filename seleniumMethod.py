import random
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys


# 主程序
def process():
    # 查询上次查到第几页了
    last_page = get_last_page_num()
    print(last_page)
    # 初始化WebDriver
    driver = webdriver.Chrome()

    # 打开目标网站
    driver.get("https://xkz.cbirc.gov.cn/jr/")
    # 使用WebDriverWait等待加载成功
    WebDriverWait(driver, 10).until(wait_web_loading)
    # 等待页面加载完成，确保页码input元素可见
    page_num_element = driver.find_element(By.ID, "ext-comp-1004")
    driver.execute_script("arguments[0].focus();", page_num_element)
    sleep()
    # 选中文本并删除
    page_num_element.send_keys(Keys.CONTROL + "a")  # 选中文本
    sleep()
    page_num_element.send_keys(Keys.BACKSPACE)  # 删除选中文本
    print('clear')
    sleep()
    page_num_element.send_keys(str(get_last_page_num() + 1))  # 设置新的页码

    page_num_element.send_keys(Keys.ENTER)
    sleep()

    # 循环3次进行翻页和数据提取
    for _ in range(10):
        res = []
        # 找到页码
        page_num = driver.find_element(By.ID, "ext-comp-1004").get_attribute("value")

        # 找到数据容器
        rows = driver.find_elements(By.CLASS_NAME, "x-grid3-row")
        for row in rows:
            # 提取每一行的数据
            data = [cell.find_element(By.TAG_NAME, "div").text for cell in
                    row.find_elements(By.TAG_NAME, "td")[1:]]  # 从1开始以跳过行号
            print(data)  # 打印或处理数据
            res.append(data)
        save_data(res, page_num)
        # 点击翻页按钮
        driver.find_element(By.ID, "ext-gen36").click()
        # 等待页面加载完成
        WebDriverWait(driver, 10).until(wait_web_loading)
        sleep()
    sleep()
    # 关闭WebDriver
    driver.quit()


def wait_web_loading(driver):
    query_result_text = driver.find_element(By.ID, "queryResult").text
    if query_result_text == '您的操作过于频繁，请五分钟后再试。':
        print("wait 5 min")
        time.sleep(400)
    try:
        # 等待reportSearch的input标签的disabled属性消失
        WebDriverWait(driver, 10).until_not(EC.element_to_be_clickable((By.ID, "reportSearch")))

        return True
    except:
        return False


def save_data(rows, page_num):
    try:
        data = {
            '序号': [row[0] for row in rows],
            '机构编码': [row[1] for row in rows],
            '流水号': [row[2] for row in rows],
            '机构名称': [row[3] for row in rows],
            '批准日期': [row[5] for row in rows],
            '发证日期': [row[6] for row in rows],
            '页码': [page_num for row in rows]
        }
        new_data = pd.DataFrame(data)

        # 检查文件是否存在，以决定是否写入header
        try:
            with open('data.csv', 'x') as f:
                new_data.to_csv(f, index=False)
        except FileExistsError:
            new_data.to_csv('data.csv', mode='a', header=False, index=False)
    except Exception as e:
        # 如果保存失败，将page_num记录到error.txt中
        with open('error.txt', 'a') as error_file:
            error_file.write(f'Failed to save page {page_num}: {e}\n')


def get_last_page_num():
    try:
        df = pd.read_csv('data.csv')
        last_page = df.iloc[-1, -1]  # 获取页码列的最后一个值
        print(last_page)
        return last_page
    except (FileNotFoundError, IndexError, pd.errors.EmptyDataError):
        last_page = 0  # 如果文件不存在或为空，则默认从第一页开始
        return last_page


def test():
    driver = webdriver.Chrome()

    # 打开目标网站
    driver.get("https://xkz.cbirc.gov.cn/jr/")
    # 等待页面加载完成，确保input元素可见
    page_num_element = driver.find_element(By.ID, "ext-comp-1004")
    driver.execute_script("arguments[0].focus();", page_num_element)
    time.sleep(1)
    # 选中文本并删除
    page_num_element.send_keys(Keys.CONTROL + "a")  # 选中文本
    time.sleep(1)
    page_num_element.send_keys(Keys.BACKSPACE)  # 删除选中文本
    print('clear')
    time.sleep(2)
    page_num_element.send_keys(str(get_last_page_num() + 1))  # 设置新的页码

    page_num_element.send_keys(Keys.ENTER)
    time.sleep(20)
    # 关闭WebDriver
    driver.quit()


def sleep():
    time.sleep(random.randint(2, 5))
