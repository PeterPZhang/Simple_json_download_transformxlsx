# -*- coding: utf-8 -*-
"""
__author__ = 'peter'
__mtime__ = '2018/12/17'
# Follow the master,become a master.
             ┏┓       ┏┓
            ┏┛┻━━━━━━━┛┻┓
            ┃    ☃      ┃
            ┃  ┳┛   ┗┳  ┃
            ┃     ┻     ┃
            ┗━┓       ┏━┛
              ┃       ┗━━━━┓
              ┃ 神兽保佑     ┣┓
              ┃　永无BUG！   ┏┛
              ┗┓┓┏━━━┳┓┏━━━┛
               ┃┫┫   ┃┫┫
               ┗┻┛   ┗┻┛
"""
import json
import os
from pathlib import Path

import requests
import xlsxwriter

os.makedirs('./image/', exist_ok=True)  # 创建存放图片路径
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 获取当前文件路径
json_file_path = os.getcwd() + '/product'  # 获取当前文件下json文件路径


# 获取当前指定路径下的所有json文件
def parse_dir(root_dir):
    """

    :param root_dir:
    :return:
    """
    path = Path(root_dir)
    all_json_file = list(path.glob('*.json'))
    parse_result = []
    for json_file in all_json_file:
        parse_result.append(str(json_file))
    return parse_result


# 生成excel文件
def generate_excel(row, col, expenses):
    """

    :param row:
    :param col:
    :param expenses:
    :return:
    """
    workbook = xlsxwriter.Workbook('./tran_data.xlsx')
    worksheet = workbook.add_worksheet()
    bold_format = workbook.add_format({'bold': True})  # 设定格式，字典中格式为指定选项
    money_format = workbook.add_format({'num_format': '$#,##0'})  # bold：加粗，num_format:数字格式
    date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
    worksheet.set_column(1, 1, 15)  # 将二行二列设置宽度为15(从0开始)
    worksheet.write('A1', '手机id', bold_format)  # 用符号标记位置，例如：A列1行
    worksheet.write('B1', '品牌', bold_format)
    worksheet.write('C1', '手机名称', bold_format)
    worksheet.write('D1', '手机图片URL', bold_format)
    worksheet.write('E1', '最高价', bold_format)
    worksheet.write('F1', '品牌id', bold_format)
    worksheet.write('G1', '存储本地图片名称', bold_format)
    for item in expenses:  # 使用write_string方法，指定数据格式写入数据
        worksheet.write_string(row, col, str(item['phone_id']))
        worksheet.write_string(row, col + 1, str(item['brand']))
        worksheet.write_string(row, col + 2, str(item['phone_name']))
        worksheet.write_string(row, col + 3, str(item['phone_img']))
        worksheet.write_string(row, col + 4, str(item['top_price']))
        worksheet.write_string(row, col + 5, str(item['brand_id']))
        worksheet.write_string(row, col + 6, str(item['img_path']))
        row += 1
    workbook.close()
    return row


# 格式化不符合标准的图片名称
def format_name(phone_name):
    """

    :param phone_name:
    :return:
    """
    full_phone_name = ''.join(phone_name.split())  # 删除空格
    brand = phone_name.split()[0]
    phone_name = ''.join(phone_name.split()[1:])
    format_phone_name = phone_name.replace('/', '、')  # 删除'/'以'、'代替（否则将认为是路径处理）
    format_full_phone_name = full_phone_name.replace('/', '、')  # 删除'/'以'、'代替（否则将认为是路径处理）
    return brand, format_phone_name, format_full_phone_name


# 根据url下载手机图片
def img_download(img_name, url):
    """

    :param img_name:
    :param url:
    :return:
    """
    headers = {
        "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36"}
    r = requests.get(url, stream=True, headers=headers)
    filename = os.path.basename(url)
    with open('./image/%s' % filename, 'wb') as f:
        print('%s 正在下载......' % img_name)
        for chunk in r.iter_content(chunk_size=32):
            f.write(chunk)
    print('%s 下载成功！' % filename)
    return filename


# 读取json
def load_json(file):
    """

    :param file:
    :return:
    """
    file = open(file, encoding='utf-8')  # 设置以utf-8解码模式读取文件，encoding参数必须设置，否则默认以gbk模式读取文件，当文件中包含中文时，会报错
    file_data = json.load(file)
    data_list = file_data['data']
    phone_list = []
    for item in data_list:
        phone_id = item['id']
        phone_name = item['name']
        phone_img = item['imgUrl']
        top_price = item['topPrice']
        brand_id = item['brandId']
        brand, phone_name, format_full_phone_name = format_name(phone_name)  # 格式化手机名称为正确格式
        file_name = img_download(format_full_phone_name, phone_img)  # 根据json中的图片url下载图片到指定位置
        phone_dict = {
            "phone_id": phone_id,
            "brand": brand,
            "phone_name": phone_name,
            "phone_img": phone_img,
            "top_price": top_price,
            "brand_id": brand_id,
            "img_path": file_name
        }
        phone_list.append(phone_dict)
    return phone_list


if __name__ == '__main__':
    ROW = 1
    COL = 0
    json_file_path_list = parse_dir(json_file_path)
    final_result = []
    for path_item in json_file_path_list:
        trans_result = load_json(path_item)
        final_result.extend(trans_result)
    generate_excel(ROW, COL, final_result)
