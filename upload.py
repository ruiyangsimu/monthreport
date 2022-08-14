# -*- coding=utf-8
from qcloud_cos import CosConfig
from qcloud_cos import CosS3Client
import sys
import logging
import argparse
import time

class COS(object):
    def __init__(self, secret_id, secret_key, region='ap-shanghai', token=None, scheme='https'):
        """
        初始化上传信息
        :param secret_id: 替换为用户的 SecretId，请登录访问管理控制台进行查看和管理，https://console.cloud.tencent.com/cam/capi
        :param secret_key: 替换为用户的 SecretKey，请登录访问管理控制台进行查看和管理，https://console.cloud.tencent.com/cam/capi
        :param region:  替换为用户的 region，已创建桶归属的region可以在控制台查看，https://console.cloud.tencent.com/cos5/bucket，COS支持的所有region列表参见https://cloud.tencent.com/document/product/436/6224
        :param token: 如果使用永久密钥不需要填入token，如果使用临时密钥需要填入，临时密钥生成和使用指引参见https://cloud.tencent.com/document/product/436/14
        :param scheme: 指定使用 http/https 协议来访问 COS，默认为 https，可不填
        """
        # 正常情况日志级别使用INFO，需要定位时可以修改为DEBUG，此时SDK会打印和服务端的通信信息
        logging.basicConfig(level=logging.INFO, stream=sys.stdout)
        self.region = region
        config = CosConfig(Region=region, SecretId=secret_id, SecretKey=secret_key, Token=token, Scheme=scheme)
        self.client = CosS3Client(config)

    def upload(self, file_name, dist_name, bucket='file-1254396400'):
        """
        上传文件
        :param bucket: 桶名称
        :param file_name: 需要上传的文件全路径或相关路径，如：./dist/test.txt
        :param dist_name: 目标文件全路径，如：./ruiyang/rj.exe
        :return:
        """
        response = self.client.upload_file(
            Bucket=bucket,
            LocalFilePath=file_name,
            Key=dist_name,
            PartSize=1,
            MAXThread=10,
            EnableMD5=False
        )
        return response

    def copy(self, dist_file, source_file, dist_bucket='file-1254396400', source_bucket='file-1254396400', source_region='ap-shanghai'):
        """
        拷贝文件
        :param dist_file: 目标文件
        :param source_file: 需要拷贝的源文件
        :param dist_bucket: 目标文件所在桶的位置
        :param source_bucket: 源文件所在桶的位置
        :param source_region: 源文件所在区域
        :return:
        """
        self.client.copy(
            Bucket=dist_bucket,
            Key=dist_file,
            CopyStatus='Replaced',
            CopySource={
                'Bucket': source_bucket,
                'Key': source_file,
                'Region': source_region
            }
        )

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='操作腾讯云COS')
    parser.add_argument('--si',
                        required=True,
                        type=str,
                        help='密钥ID')
    parser.add_argument('--sk',
                        required=True,
                        type=str,
                        help='密钥key')
    parser.add_argument('--sf',
                        required=False,
                        type=str,
                        help='上传/拷贝文件全路径，如：dist/main.exe')
    parser.add_argument('--df',
                        required=False,
                        type=str,
                        help='目标文件全路径，如：ruiyang/ruiyang.exe')
    args = parser.parse_args()
    cos = COS(secret_id=args.si, secret_key=args.sk)
    cos.upload(file_name=args.sf, dist_name=args.df)
    now_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    cos.copy(source_file=args.df, dist_file="./ruiyang/ruiyang_" + now_time + ".exe",
             source_bucket='file-1254396400')

