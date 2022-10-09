from zipfile import ZipFile
import os
import shutil


def compress_file(zipfilename, dirname):
    """
    将文件夹压缩压缩包中
    :param zipfilename: 压缩包名称
    :param dirname: 需要打包的目录
    :return:
    """
    if os.path.isfile(dirname):
        with ZipFile(zipfilename, 'w') as z:
            z.write(dirname)
    else:
        with ZipFile(zipfilename, 'w') as z:
            for root, dirs, files in os.walk(dirname):
                for single_file in files:
                    if single_file != zipfilename:
                        filepath = os.path.join(root, single_file)
                        z.write(filepath)


def addfile(zipfilename, dirname):
    if os.path.isfile(dirname):
        with ZipFile(zipfilename, 'a') as z:
            z.write(dirname)
    else:
        with ZipFile(zipfilename, 'a') as z:
            for root, dirs, files in os.walk(dirname):
                for single_file in files:
                    if single_file != zipfilename:
                        filepath = os.path.join(root, single_file)
                        z.write(filepath)


if __name__ == '__main__':
    if os.path.isfile("ruiyang.exe"):
        os.remove("ruiyang.exe")
    if os.path.isfile("./dist/ruiyang.zip"):
        os.remove("./dist/ruiyang.zip")
    compress_file('./dist/ruiyang.zip', './data')
    addfile('./dist/ruiyang.zip', './config')
    current_path = os.path.abspath(__file__)
    dir_name = os.path.dirname(current_path)
    shutil.copy(dir_name+'\\dist\\ruiyang.exe', dir_name+'\\')
    addfile('./dist/ruiyang.zip', './ruiyang.exe')
