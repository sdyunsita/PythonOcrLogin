# -*- coding: UTF-8 -*-
# 这个文件是写excel表格的，跟Selenium和OCR没关系
import openpyxl
from openpyxl.drawing.image import Image
import shutil
import datetime


def generate_report(domain_list, name_list, cishu_list, yuanip, yuandiyu, gongjicishu, other_list, ipv6c):
    print('开始生成报告......')
    newfilename = './xx学院周报2022.06.17.xlsx'
    shutil.copy('周报模板.xlsx', newfilename)  # 复制报告模板以指定路径+名称存放
    workbook = openpyxl.load_workbook(newfilename)
    sht = workbook.worksheets[0]
    huizongxinxi = f"1、本周防护资产235个（包括 http/https/ipv4/ipv6），0个资产变更，IPV6 支持度评分{ipv6c[4]},共发现0个漏洞，新业务上线0个，暗链接0个，挖矿木马0个;\n2、本周未发现挖矿活动情况，处于可控并动态清零状态。"
    sht['B7'] = huizongxinxi
    start_num = 16
    for item in domain_list:
        sht['F' + str(start_num)] = item
        start_num += 1
    start_num = 16
    for item in name_list:
        sht['B' + str(start_num)] = item
        start_num += 1
    start_num = 16
    for item in cishu_list:
        sht['I' + str(start_num)] = item
        start_num += 1

    start_num = 24
    for item in yuanip:
        sht['B' + str(start_num)] = item
        start_num += 1
    start_num = 24
    for item in yuandiyu:
        sht['D' + str(start_num)] = item
        start_num += 1
    start_num = 24
    for item in gongjicishu:
        sht['G' + str(start_num)] = item
        start_num += 1
    imgPath = r'./pics/qushi.png'
    img = Image(imgPath)
    img.width, img.height = (520, 150)
    sht.add_image(img, 'B32')

    imgPath = r'./pics/mudiip.png'
    img = Image(imgPath)
    img.width, img.height = (170, 200)
    sht.add_image(img, 'B40')

    imgPath = r'./pics/gongjileixing.png'
    img = Image(imgPath)
    img.width, img.height = (170, 200)
    sht.add_image(img, 'F40')
    sht['B53'] = f'最近七天共防护攻击{other_list[0]}，其中{other_list[1]}次数最多{other_list[2]}次。本周所有攻击行为均被阻断，防护资产合计235个（包括 http/https/ipv4/ipv6），非重保时期网盾防护策略为通用防护策略，上述攻击次数较多的地址可在校内防火墙予以封禁，若有需要防护的资产增加，请及时联系我方工作人员。'
    d = datetime.datetime.today().strftime('%Y.%m.%d')
    sht['H59'] = d
    
    sht['B64'] = ipv6c[0]
    sht['C64'] = ipv6c[1]
    sht['F64'] = ipv6c[2]
    sht['G64'] = ipv6c[3]
    sht['H64'] = ipv6c[4]
    sht['I64'] = ipv6c[5]
    workbook.save(newfilename)
