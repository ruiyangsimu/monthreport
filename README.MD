# 安装开发和编译环境
1. 安装python 3.7.6版本
2. `pip install virtualenv`安装用于创建虚拟环境的包
3. 进入项目目录后执行`virtualenv venv`命令后在本项目venv目录下生成新的python虚拟环境
4. 在window的powershell命令行工具中执行`.\venv\Scripts\activate.ps1`以激活新的python虚拟环境
5. 在window的powershell中执行`.\bin\install.ps1`脚本来安装项目所需依赖
6. 因为我使用的是pycharm进行开发，所以在pycharm中设置python编译器路径为新创建的虚拟环境（File -> Settings -> Python Interpreter[路径为：.\venv\Scripts\python.exe]）

# 本地打包
1. 首先需要在powershell中进入本项目目录使用`.\venv\Scripts\activate.ps1`命令激活
2. 执行`.\bin\build.ps1`,生成的目标文件为.\dist\ruiyang.exe

# 界面开发
1. powershell中进入本项目目录，执行`.\venv\Lib\site-packages\qt5_applications\Qt\bin\designer.exe`打开QT Designer设计软件
2. 打开.\ui\main_window.ui文件后进行编辑
3. 编辑完成后在powershell中执行`.\bin\genui.ps1`,生成main_window_ui.py
4. 运行app.py

----
参考
> https://blog.karatos.in/a?ID=00950-d548251c-9168-43dd-86e2-e652770e4769#%E5%BB%BA%E7%AB%8B%20images.qrc