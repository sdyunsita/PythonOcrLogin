import ddddocr
from PIL import Image
from selenium.webdriver.common.by import By
import base64
import json
import requests
import time
import urllib
from msedge.selenium_tools import Edge, EdgeOptions



driver = 'MicrosoftWebDriver.exe'
options = EdgeOptions()
options.use_chromium = True
options.add_experimental_option('excludeSwitches', ['enable-logging'])
# options.add_argument('--ignore-certificate-errors --headless')    # 隐藏浏览器窗口
options.add_argument('--ignore-certificate-errors')
browser = Edge(options=options, executable_path=driver)
browser.implicitly_wait(10)
browser.get('https://x.x.x.x/user/login')
browser.set_window_size(width=1265,height=1372) # 宽点好看
# print(browser.page_source)
input_username = browser.find_element(By.ID, 'username')
input_username.send_keys('你的账号')
input_password = browser.find_element(By.ID, 'password')
input_password.send_keys('你的密码')

# 获取验证码（后期自动识别）
code_img = browser.find_element(By.XPATH,
                                '/html/body/div/div/div/div[2]/div[1]/div[2]/div/form/div[1]/div[2]/div/div/div[3]/div/div/span/span/span[2]/img')
img = code_img.get_attribute('src')
# print(img.split(',')[1])
imgdata = base64.b64decode(img.split(',')[1])
file = open('./pics/code.png', 'wb')
file.write(imgdata)
file.close()
# 自动识别验证码
ocr = ddddocr.DdddOcr()
with open(r'./pics/code.png', 'rb') as f:
    img_bytes = f.read()
code = ocr.classification(img_bytes)
input_code = browser.find_element(By.ID, 'code')
input_code.send_keys(code)
# 点击登录
login_button = browser.find_element(By.XPATH,
                                    '//*[@id="root"]/div/div/div[2]/div[1]/div[2]/div/form/div[2]/div/div/span')
login_button.click()
time.sleep(3)

# 转到日志报表页面
browser.get('https://x.x.x.x/logreport/loganalysis')
time.sleep(2)

# 设置日期并查询最近7天
date1 = browser.find_element(By.XPATH,
                             '/html/body/div[1]/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/form/div/div[1]/div/div[2]/div/div/div/div[1]/input')
date1.click()
time.sleep(1)
date2 = browser.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div[2]/div[2]/ul/li[3]/span')
date2.click()
chaxun = browser.find_element(By.XPATH,
                              '/html/body/div/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/form/div/div[14]/div/div[2]/div/div/div/div[1]/div/div[1]')
chaxun.click()
time.sleep(2)

# 转到攻击目的tab
gongjimudi_tab = browser.find_elements(By.CLASS_NAME, 'ant-tabs-tab')[2]
gongjimudi_tab.click()
time.sleep(2)

domain_list = []
for i in range(1, 6):
    text = browser.find_element(By.XPATH,
                                f'/html/body/div[1]/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/div[2]/div[2]/div/div[3]/div/div/div/div[2]/div/div/div/div/div/table/tbody/tr[{i}]/td[5]/span').text
    if '127.0.0.1' in text:
        text = browser.find_element(By.XPATH,
                                    f'/html/body/div[1]/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/div[2]/div[2]/div/div[3]/div/div/div/div[2]/div/div/div/div/div/table/tbody/tr[{i}]/td[3]').text
    domain_list.append(text)
    print(text)

# 转到综合统计分析tab
zonghefenxi_tab = browser.find_elements(By.CLASS_NAME, 'ant-tabs-tab')[3]
zonghefenxi_tab.click()
time.sleep(12)


def join_images(png1, png2, size=0, output='./pics/result.png'):
    """
    图片拼接
    :param png1: 图片1
    :param png2: 图片2
    :param size: 两个图片重叠的距离
    :param output: 输出的图片文件
    :return:
    """
    # 图片拼接
    img1, img2 = Image.open(png1), Image.open(png2)
    size1, size2 = img1.size, img2.size  # 获取两张图片的大小
    joint = Image.new('RGB', (size1[0], size1[1] + size2[1] - size))  # 创建一个空白图片
    # 设置两张图片要放置的初始位置
    loc1, loc2 = (0, 0), (0, size1[1] - size)
    # 分别放置图片
    joint.paste(img2, loc2)
    joint.paste(img1, loc1)
    # 保存结果
    joint.save(output)


JS = {
    '滚动到页尾': "window.scroll({top:document.body.clientHeight,left:0,behavior:'auto'});",
    '滚动到': "window.scroll({top:%d,left:0,behavior:'auto'});",
}
# 获取body大小
body_h = int(
    browser.find_element(By.XPATH, '/html/body/div/div/div/div/section/section/main/section/main/div').size.get(
        'height'))
browser.save_screenshot('./pics/result.png')
# 计算当前页面截图的高度
# （使用driver.get_window_size()也可以获取高度，但有误差，推荐使用图片高度计算）
current_h = Image.open('./pics/result.png').size[1]
image_list = ['./pics/result.png']  # 储存截取到的图片路径

for i in range(1, int(body_h / current_h)):
    # 1. 滚动到指定锚点
    browser.execute_script(JS['滚动到'] % (current_h * i))
    # 2. 截图
    browser.save_screenshot(f'./pics/test_{i}.png')
    join_images('./pics/result.png', f'./pics/test_{i}.png')
# 处理最后一张图
browser.execute_script(JS['滚动到页尾'])
browser.save_screenshot('./pics/test_end.png')
# 拼接图片
join_images('./pics/result.png', './pics/test_end.png', size=current_h - int(body_h % current_h) - 90)


def jietu(ele, name):
    print(ele.location)
    print(ele.size)
    left = ele.location['x']
    top = ele.location['y']

    right = ele.location['x'] + ele.size['width']
    bottom = ele.location['y'] + ele.size['height']

    photo = Image.open('./pics/result.png')
    photo = photo.crop((left, top, right, bottom))
    photo.save(f'./pics/{name}.png')


# 目的IP截图
mudiIP = browser.find_element(By.XPATH,
                              '/html/body/div[1]/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/div[2]/div[2]/div/div[4]/div/div/div/div[1]/div[3]')
jietu(mudiIP, name='mudiip')

# 攻击类型截图
gongjileixing = browser.find_element(By.XPATH,
                                     '/html/body/div[1]/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/div[2]/div[2]/div/div[4]/div/div/div/div[1]/div[4]')
jietu(gongjileixing, name='gongjileixing')

# 时间分布趋势截图
qushi = browser.find_element(By.XPATH,
                             '/html/body/div[1]/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/div[2]/div[2]/div/div[4]/div/div/div/div[4]')
jietu(qushi, name='qushi')

# 获取网站名称及攻击次数
mudiip_more = browser.find_element(By.XPATH,
                                   '/html/body/div[1]/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/div[2]/div[2]/div/div[4]/div/div/div/div[1]/div[3]/div/div/div[1]/div/div[2]/button')
browser.execute_script("arguments[0].click();", mudiip_more)
time.sleep(2)
name_list = []
cishu_list = []
for i in range(1, 6):
    text_name = browser.find_element(By.XPATH,
                                     f'/html/body/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div/div/table/tbody/tr[{i}]/td[3]').text
    name_list.append(text_name)
    text_cishu = browser.find_element(By.XPATH,
                                      f'/html/body/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div/div/table/tbody/tr[{i}]/td[4]').text
    cishu_list.append(text_cishu)
    print(text_name, text_cishu)
# 关闭窗口
browser.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div/div[2]/button').click()


# 获取攻击类型统计数据
other_list = []
other_0 = browser.find_element(By.XPATH,'/html/body/div[1]/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/div[2]/div[2]/div/div[4]/div/div/div/div[1]/div[4]/div/div/div[2]/div/div/div/h4/div/div/div[2]/font').text
other_list.append(other_0)
gonejileixing_more = browser.find_element(By.XPATH, '/html/body/div/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/div[2]/div[2]/div/div[4]/div/div/div/div[1]/div[4]/div/div/div[1]/div/div[2]/button')
browser.execute_script("arguments[0].click();", gonejileixing_more)
time.sleep(2)
other_1 = browser.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div/div/table/tbody/tr[1]/td[2]').text
other_list.append(other_1)
other_2 = browser.find_element(By.XPATH,'/html/body/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div/div/table/tbody/tr[1]/td[3]').text
other_list.append(other_2)
# 关闭窗口
browser.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div/div[2]/button').click()


# 获取攻击源IP信息
yuanip_more = browser.find_element(By.XPATH,
                                   '/html/body/div[1]/div/div/div/section/section/main/section/main/div/div[2]/div/div/div/div[2]/div[2]/div/div[4]/div/div/div/div[1]/div[1]/div/div/div[1]/div/div[2]/button')
browser.execute_script("arguments[0].click();", yuanip_more)
time.sleep(2)

yuanip = []
yuandiyu = []
gongjicishu = []

for i in range(1, 6):
    yuanip_text = browser.find_element(By.XPATH,
                                       f'/html/body/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div/div/table/tbody/tr[{i}]/td[2]').text
    yuandiyu_text = browser.find_element(By.XPATH,
                                         f'/html/body/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div/div/table/tbody/tr[{i}]/td[3]').text
    gongjicishu_text = browser.find_element(By.XPATH,
                                            f'/html/body/div[3]/div/div[2]/div/div[2]/div[2]/div/div/div/div/div/div/table/tbody/tr[{i}]/td[4]').text
    yuanip.append(yuanip_text)
    yuandiyu.append(yuandiyu_text)
    gongjicishu.append(gongjicishu_text)
    print(yuanip_text, yuandiyu_text, gongjicishu_text)

print(domain_list, name_list, cishu_list, yuanip, yuandiyu, gongjicishu, other_list)



# 获取IPv6支持度
browser.get("http://ipv6c.cn/2021.do?cer2Date=2022%2F05%2F13&inputProv=%E5%B1%B1%E4%B8%9C&selectType=2&lineType=1")
school_name = browser.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/table/tbody/tr[102]/td[3]").text
school_url = browser.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/table/tbody/tr[102]/td[4]/a").text
erji_zhichidu = browser.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/table/tbody/tr[102]/td[7]").text
sabji_zhichidu = browser.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/table/tbody/tr[102]/td[8]").text
pingfen = browser.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/table/tbody/tr[102]/td[9]").text
huoyue = browser.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div/table/tbody/tr[102]/td[10]").text

ipv6c = [school_name,school_url,erji_zhichidu,sabji_zhichidu,pingfen,huoyue]

browser.close()
import excel
excel.generate_report(domain_list, name_list, cishu_list, yuanip, yuandiyu, gongjicishu, other_list, ipv6c)

