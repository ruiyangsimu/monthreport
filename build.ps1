# 创建虚拟环境，首次下载项目时打开注释执行，创建属于本项目的python环境，防止环境污染
# virtualenv venv
# 激活虚拟环境
.\venv\Scripts\activate.ps1
# 安装包
pip install -r .\requirements.txt
pyinstaller --clean --win-private-assemblies --key 0123456789 -F main.py