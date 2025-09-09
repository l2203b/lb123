from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service  # 用于配置驱动路径
import time
import pandas as pd  # 处理Excel文件

class ExcelEdgeAutomation:
    def __init__(self):
        # 配置Edge浏览器选项，连接远程调试模式
        edge_options = Options()
        edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        
        # 配置驱动路径（确保msedgedriver.exe与脚本同目录）
        driver_path = "./msedgedriver.exe"  # 相对路径，同目录下
        # 若驱动在其他位置，用绝对路径，例如：
        # driver_path = "D:/工具/msedgedriver.exe"
        
        service = Service(executable_path=driver_path)
        self.driver = webdriver.Edge(service=service, options=edge_options)
        self.wait = WebDriverWait(self.driver, 15)  # 延长等待时间至15秒，确保元素加载

    def query_phone(self, id_value, input_xpath, query_btn_xpath, phone_xpath):
        """根据ID查询手机号：先输入ID，再点击查询按钮，最后获取手机号"""
        try:
            # 1. 定位并清空输入框，输入ID（处理长数字）
            input_box = self.wait.until(EC.element_to_be_clickable((By.XPATH, input_xpath)))
            input_box.clear()  # 清空原有内容
            
            # 处理长数字：避免科学计数法，强制以完整数字字符串形式输入
            if isinstance(id_value, (int, float)):
                id_str = f"{id_value:.0f}"  # 格式化长数字为完整字符串
            else:
                id_str = str(id_value)  # 本身是字符串则直接转换
            
            input_box.send_keys(id_str)  # 输入处理后的UID
            time.sleep(0.5)  # 短暂等待等待输入完成

            # 2. 定位并点击“查询”按钮
            query_btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, query_btn_xpath)))
            query_btn.click()  # 点击查询
            time.sleep(2)  # 等待查询结果加载（根据页面响应速度调整）

            # 3. 定位并返回手机号结果（使用你提供的XPath）
            phone_element = self.wait.until(EC.presence_of_element_located(
                (By.XPATH, phone_xpath)  # 使用指定的手机号XPath
            ))
            return phone_element.text.strip()  # 返回手机号
        
        except Exception as e:
            print(f"查询ID={id_value}失败: {str(e)}")
            return "查询失败"  # 失败标记

    def process_excel(self, excel_path, input_xpath, query_btn_xpath, phone_xpath, id_column="UID", phone_column="number"):
        """处理Excel：读取UID列，批量查询后写入结果"""
        # 读取Excel文件时，强制将UID列按字符串类型加载（避免科学计数法）
        df = pd.read_excel(excel_path, dtype={id_column: str})
        
        # 检查UID列是否存在
        if id_column not in df.columns:
            print(f"错误：Excel中未找到'{id_column}'列，请检查表头")
            return
        
        # 若结果列不存在，自动创建
        if phone_column not in df.columns:
            df[phone_column] = ""
        
        # 遍历每一行进行查询
        total = len(df)
        for index, row in df.iterrows():
            current_id = row[id_column]
            if pd.isna(current_id) or current_id.strip() == "":  # 跳过空值或空白
                continue
            
            # 显示进度并查询
            print(f"处理进度: {index+1}/{total}，UID={current_id}")
            phone = self.query_phone(current_id, input_xpath, query_btn_xpath, phone_xpath)
            df.at[index, phone_column] = phone  # 写入结果
        
        # 保存结果到原文件
        df.to_excel(excel_path, index=False)
        print(f"\n所有操作完成，结果已保存至：{excel_path}")

    def close(self):
        """关闭浏览器驱动"""
        self.driver.quit()

if __name__ == "__main__":
    # 配置参数（根据你的信息设置）
    EXCEL_FILE_PATH = "D:/桌面/工作簿1.xlsx"  # Excel路径
    INPUT_XPATH = "//*[@id='app']/div/div[2]/section/div/div/div[1]/form/section/div[9]/div/div/div/input"  # UID输入框XPath
    QUERY_BTN_XPATH = "//*[@id='app']/div/div[2]/section/div/div/div[1]/form/div[2]/button[2]/span"  # 查询按钮XPath
    PHONE_XPATH = "//*[@id='app']/div/div[2]/section/div/div/div[2]/div[2]/div[1]/div[3]/table/tbody/tr/td[3]/div"  # 手机号结果XPath
    ID_COLUMN_NAME = "UID"  # A列头
    PHONE_COLUMN_NAME = "number"  # B列头

    try:
        automation = ExcelEdgeAutomation()
        automation.process_excel(
            excel_path=EXCEL_FILE_PATH,
            input_xpath=INPUT_XPATH,
            query_btn_xpath=QUERY_BTN_XPATH,
            phone_xpath=PHONE_XPATH,
            id_column=ID_COLUMN_NAME,
            phone_column=PHONE_COLUMN_NAME
        )
        input("按Enter键关闭脚本...")  # 暂停等待查看结果
    except Exception as e:
        print(f"执行出错: {str(e)}")
    finally:
        if 'automation' in locals():
            automation.close()  # 确保关闭浏览器
    
