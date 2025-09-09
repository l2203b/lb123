一、前期准备工作
1. 安装必要软件
- Python 环境：安装 Python 3.7 及以上版本（推荐 3.9）
  - 下载地址：https://www.python.org/downloads/
  - 安装时勾选 "Add Python to PATH"
- Edge 浏览器：确保已安装 Microsoft Edge 浏览器
  - 默认安装路径：C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
2. 安装所需 Python 库
打开命令提示符（Win+R 输入cmd），输入以下命令，Enter：
pip install selenium pandas openpyxl

- selenium：用于浏览器自动化操作
- pandas：用于处理 Excel 文件
- openpyxl：支持 pandas 读写 Excel 文件
3. 下载浏览器驱动
- 查看 Edge 浏览器版本：打开 Edge → 设置 → 关于 Microsoft Edge
- 下载对应版本的 Edge 驱动（msedgedriver）    例：版本 139.0.3405.125 (正式版本) (64 位)
  - 下载地址：https://developer.microsoft.com/zh-cn/microsoft-edge/tools/webdriver/
- 将下载的msedgedriver.exe保存到与脚本相同的文件夹   建议保存到D:\桌面\手机号查询脚本
4. 准备 Excel 文件
- 创建 Excel 文件（如工作簿1.xlsx），保存路径建议为：D:/桌面/工作簿1.xlsx
- 表格格式要求：
  - 包含列头为 "UID" 的列（存放需要查询的 UID）
  - 脚本会自动创建 "number" 列存放查询结果
  - 保证UID为纯文本格式，避免错误


二、脚本配置
1. 将提供的 Python 代码保存为phone_query.py（与msedgedriver.exe同目录）
2. 脚本中的配置参数：Excel文件路径和XPath（即元素在网页结构中的位置，该脚本需要UID输入框、查询按钮、手机号结果的XPath）（位于if name == "__main__":下方）：
3. python
4. 运行
EXCEL_FILE_PATH = "D:/桌面/工作簿1.xlsx"  # 你的Excel文件路径
INPUT_XPATH = "//*[@id='app']/div/div[2]/section/div/div/div[1]/form/section/div[9]/div/div/div/input"  # UID输入框XPath
QUERY_BTN_XPATH = "//*[@id='app']/div/div[2]/section/div/div/div[1]/form/div[2]/button[2]/span"  # 查询按钮XPath
PHONE_XPATH = "//*[@id='app']/div/div[2]/section/div/div/div[2]/div[2]/div[1]/div[3]/table/tbody/tr/td[3]/div"  # 手机号结果XPath

三、运行脚本的步骤
步骤 1：启动带调试模式的 Edge 浏览器
1. 打开命令提示符（Win+R 输入cmd）
2. 输入以下命令并回车：
3. cmd
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\EdgeDebugProfile"

  - 注意：如果 Edge 安装路径不同，请修改引号中的路径
  - 首次运行会创建C:\EdgeDebugProfile文件夹，用于存放独立的浏览器配置
1. 在启动的浏览器中：
  - 手动打开岚图会员积分管理后台，清除日期等筛选条件
步骤 2：运行脚本
1. 打开新的命令提示符窗口
2. 导航到脚本所在目录（例如）：
3. cmd
cd D:\桌面\手机号查询脚本  # 替换为你的脚本存放目录

1. 执行脚本：
2. cmd
python phone_query.py

步骤 3：监控运行过程
- 脚本运行时会在命令提示符窗口显示进度：处理进度: X/Y，UID=xxx
- 浏览器会自动进行操作：输入 UID → 点击查询 → 记录结果
- 若出现错误，会显示具体失败信息（如查询ID=xxx失败: ...）
步骤 4：完成操作
- 脚本运行结束后，会显示：所有操作完成，结果已保存至：D:/桌面/工作簿1.xlsx
- 按 Enter 键关闭脚本
- 查看 Excel 文件，"number" 列已填充查询到的手机号
- 
四、常见问题及解决方法
1. "无法连接到浏览器" 错误
  - 检查是否已用调试模式启动浏览器
  - 确认命令中的端口号9222与脚本中一致
2. "元素未找到" 错误
  - 检查 XPath 是否正确（可通过浏览器开发者工具验证）
  - 确保浏览器停留在正确的查询页面
3. Excel 文件无法打开
  - 确保 Excel 文件已关闭
  - 检查文件路径是否正确
4. 驱动相关错误
  - 确认msedgedriver.exe与脚本在同一目录
  - 检查驱动版本与 Edge 浏览器版本是否匹配
  - 
五、注意事项
1. 脚本运行期间，请勿手动操作浏览器，请勿打开脚本正在读写的 Excel 文件（可以打开其他excel）
2. 平均100个查询会出现3个错误，即前后行手机号重复，可以条件格式突出重复项修改，若出现大量重复，是由于网络原因，可以重新运行一次脚本流程
3. 若网络较慢，可适当延长等待时间（修改time.sleep()和WebDriverWait的参数）

代码如下：
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
    

