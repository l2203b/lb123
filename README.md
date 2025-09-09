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

