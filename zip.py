from zipfile import ZipFile
import os


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


if __name__ == '__main__':
    compress_file('a.zip', './data')
