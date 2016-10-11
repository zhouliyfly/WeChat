# coding:utf-8

from selenium import webdriver
from bs4 import BeautifulSoup
import urllib.parse
import urllib.request
import json
import requests
import re
import time
import xlwt
import ssl


# -----------------------------------------------------------------------
# 时间转换函数
# -----------------------------------------------------------------------

# ----------------------
# 标准时间转换为时间戳
# @ timeA : 包含标准时间字符串 "2016-08-09"
# @ 返回： 时间戳
# ----------------------
def time_to_stramp(timeA):
    astr = timeA
    # 取出字符串中的标准时间
    time_list = re.findall(r"\d+", astr)
    str_time = '-'.join(time_list)
    timeA = time.strptime(str_time, "%Y-%m-%d")
    timeStamp = int(time.mktime(timeA))
    return timeStamp


# ----------------------
# 时间戳转换为标准时间
# @ create_time: 时间戳
# @ 返回：标准时间格式 "2016-08-09"
# ----------------------
def stamp_to_time(create_time):
    timeStamp = create_time
    timeArray = time.localtime(timeStamp)

    return time.strftime("%Y-%m-%d %H:%M:%S", timeArray)


# -----------------------------------------------------------------------
# selenium自动化函数
# -----------------------------------------------------------------------

# ----------------------
# 浏览器选择
# PhantomJS/Chrome/Firefox 三种都可以，默认选择PhantomJS

# @ browser_num：浏览器选择
#   0: PhantomJS -------default
#   1: Chrome 
#   2: Firefox
# @ 返回值：浏览器对象，选择错误返回None
# ----------------------
def choose_browser(browser_num=0):
    if (browser_num == 0):
        return webdriver.PhantomJS()
    elif (browser_num == 1):
        return webdriver.Chrome()
    elif (browser_num == 2):
        return webdriver.Firefox()
    else:
        # print('browser choose error!')
        return None


# ----------------------
# 模拟登陆微信公众号
# 自动登录返回二维码验证图片，微信扫码之后进入公众号页面，最后跳转到用户管理页面
# @ 返回值：webdriver打开的浏览器对象
# ----------------------
def longin_weixingzh(browser_name=1):
    # 测试阶段使用chrom浏览器，便于观察网页登录和跳转情况
    driver = choose_browser(browser_name)
    url = 'https://mp.weixin.qq.com'
    driver.get(url)
    elem_account = driver.find_element_by_name('account')
    elem_pwd = driver.find_element_by_name('password')
    elem_login = driver.find_element_by_id('loginBt')

    # 自动填写用户名和密码
    elem_account.send_keys(u'yishidayi')
    elem_pwd.send_keys('wddzl&2017')
    elem_login.click()
    # 最大化窗口
    driver.maximize_window()
    # print(driver.current_url)

    #
    # TODO:返回二维码验证图片
    #

    #
    # TODO:判断页面是否跳转（成功登录公众号）
    # 可以通过判断当前url是否变化实现，或者判断url中是否有'token=****'
    #

    # 等待手动扫码
    print('loadeding<<<<<<<<<<<<<<<<<')
    time.sleep(25)

    # 跳转到用户管理页面
    elem_btn = driver.find_element_by_css_selector('#menuBar > dl:nth-child(2) > dd:nth-child(3) > a')
    elem_btn.click()
    time.sleep(1)

    return driver


# -----------------------------------------------------------------------
# 网页请求函数
# -----------------------------------------------------------------------

# ----------------------
# 获取当前网页url和cookies
# @ driver：webdriver打开的浏览器对象
# @ 返回值：网页信息（url/cookies/token）
# ----------------------
def get_current_pageinfo(driver):
    cookies_list = [item['name'] + "=" + item["value"] for item in driver.get_cookies()]
    current_page_cookies = ';'.join(item for item in cookies_list)
    current_page_url = driver.current_url
    current_token = current_page_url.split('=')[-1]

    page_info = {
        'url': current_page_url,
        'cookies': current_page_cookies,
        'token': current_token
    }
    return page_info


# ----------------------
# 使用requests请求url
# @ url：请求的url地址
# @ headers：网址header信息，包含cookies；这里加上这个参数实现免登陆
# @ 返回值：请求得到的页面信息
# ----------------------
def use_requests_for_url(url, headers=None):
    wb_data = requests.get(url, headers=headers)

    return wb_data


# ----------------------
# 使用BeautifulSoup解析网页
# @ web_data：请求得到的网页信息
# ----------------------
def parase_by_beautifulsoup(wb_data):
    soup = BeautifulSoup(wb_data.text, 'lxml')

    return soup


# -----------------------------------------------------------------------
# 网页处理函数
# -----------------------------------------------------------------------

# ----------------------
# 解析网页：使用requests库完成请求，BeautifulSoup库解析网页
# @ url：指定网页地址
# @ 返回值：BeautifulSoup解析之后的结果
# ----------------------
def html_parase(url, headers=None):
    wb_data = use_requests_for_url(url, headers)
    soup = parase_by_beautifulsoup(wb_data)

    return soup


# -----------------------------------------------------------------------
# 获取用户信息函数
# -----------------------------------------------------------------------

# ----------------------
# 获取第一个用户信息
# @ token：当前服务器分配的token
# @ cur_cookies：当前浏览器cookies
# @ Soup_page: BeautifulSoup解析的页面信息
# @ 返回值：字典列表（仅包含一个字典），字典记录了第一个用户的信息
# ----------------------
def get_latest_user_info(token, cur_cookies, soup_page):
    # 获取第一个用户open_id
    first_user_info_list = soup_page.select(
        '#userGroups > tr:nth-of-type(1) > td.table_cell.user > div > a.remark_name')
    first_open_id = first_user_info_list[0]['data-fakeid']
    print('first_open_id=' + first_open_id)

    # 获取第一个用户详细信息
    req = get_more_info(token, cur_cookies, first_open_id)
    # 解析json格式信息
    dates = json.loads(req)
    print(dates)

    # 返回列表
    return dates['user_list']['user_info_list']


# ----------------------
# 获取指定用户详细信息
# @ token：当前服务器分配的token
# @ cur_cookies：当前浏览器cookies
# @ user_openid：待获取信息的用户openid
# @ 返回值：用户信息
# ----------------------
def get_more_info(token, cur_cookies, user_openid):
    headers = {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Encoding': 'gzip, deflate',
        'Host': 'mp.weixin.qq.com',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Referer': 'https://mp.weixin.qq.com/cgi-bin/user_tag?action=get_all_data&lang=zh_CN&token=451422775',
        'Connection': 'keep-alive',
        'Cookie': cur_cookies
    }

    url = 'https://mp.weixin.qq.com/cgi-bin/user_tag?action=get_fans_info'
    pdata = {
        'token': token,
        'lang': 'zh_CN',
        'f': 'json',
        'ajax': '1',
        'random': '0.8479345601126673',
        'user_openid': user_openid
    }
    # ----------------------python3------------------------
    tmp_pdata = urllib.parse.urlencode(pdata).encode('utf-8')
    req = urllib.request.Request(url, tmp_pdata, headers)
    # 处理网址证书问题
    context = ssl._create_unverified_context()
    r = urllib.request.urlopen(req, context=context)


    # ----------------------python2------------------------
    # need: import urllib2
    # tmp_pdata = urllib.urlencode(pdata).encode('utf-8')
    # req = urllib2.Request(url, tmp_pdata, headers)
    # r= urllib2.urlopen(req)

    user_more_info = r.read().decode('utf-8')

    print(user_more_info)
    return user_more_info


# ----------------------
# 获取所有用户简单信息（除第一个用户，前面已经得到）
# @ cur_url：当前页面url
# @ token：token值（字符串形式）
# @ cur_cookies：当前浏览器cookies
# @ fans_num：用户总数。理论上应该减去1（第一个用户信息已经获取）。不过由于不影响结果，这里简单处理
# @ first_user_info：第一个用户信息
# @ 返回值：用户openid列表
# ----------------------
def get_users_simple_info(cur_url, token, cur_cookies, fans_num, first_user_info):
    # 构建目标url，通过向此url发送请求获取用户简单信息
    url_parameters = {
        'action': 'get_user_list',
        'groupid': '-2',
        'offset': '0',
        'backfoward': '1',
        'lang': 'zh_CN',
        'f': 'json',
        'ajax': '1',
        'random': '0.7434098680425287',
        'token': token,
        'begin_openid': first_user_info['user_openid'],
        'begin_create_time': str(first_user_info['user_create_time']),
        'limit': fans_num
    }

    headers = {
        'cookie': cur_cookies
    }

    base_url = cur_url.split('?')[0]

    # 转换网页参数的格式（以&分隔的键值对）
    parameters = [ke + '=' + va for (ke, va) in url_parameters.items()]
    parameters_str = "&".join(item for item in parameters)
    obj_url = base_url + '?' + parameters_str
    print(obj_url)

    # 请求用户数据，用json解析
    web_data = requests.get(obj_url, headers=headers)
    user_simple_datas = json.loads(web_data.text)

    # 获得所有用户user_openid
    users_openid_list = [user_info['user_openid'] for user_info in user_simple_datas['user_list']['user_info_list']]

    # 插入第一个用户的open_id到列表第一个位置，构成完整的open_id列表
    users_openid_list.insert(0, first_user_info['user_openid'])

    return users_openid_list


# ----------------------
# 获取用户总数
# @ soup_page：BeautifulSoup解析的网页数据
# @ 返回值：用户总人数（字符串）
# ----------------------
def get_numbers_of_fans(soup_page):
    num_txt = soup_page.find('em', {'class': 'num'}).get_text()

    # 用正则表达式获取数字
    return re.findall(r"\d+", num_txt)[0]


# -----------------------------------------------------------------------
# Python数据导出到Excel函数
# -----------------------------------------------------------------------

# ----------------------
# 导出用户数据到Excel
# @ users_info_list：用户数据列表
# ----------------------
def export_to_excel(users_info_list, export_path=None):
    # 设置Excel文件utf-8编码，写入方式为覆盖
    workbook = xlwt.Workbook(encoding='utf-8')
    booksheet = workbook.add_sheet('用户数据', cell_overwrite_ok=True)
    # 写入标题栏
    for i, elem in enumerate(users_info_list[0]):
        booksheet.write(0, i, elem)

    # 设置列宽 256是基本单位，其后是字符数，这里要再斟酌
    #booksheet.col(0).width = 256 * 40

    # 写入数据
    for i, user in enumerate(users_info_list):
        for j, elem in enumerate(user.keys()):
            if (elem == 'user_create_time'):
                booksheet.write(i + 1, j, stamp_to_time(user[elem]))
            else:
                booksheet.write(i + 1, j, user[elem])

    # 确认文件名
    # 如果没有指定文件名，使用当前时间作为文件名
    if(export_path == None):
        file_name = time.strftime("%Y%m%d%H%M%S", time.localtime(time.time())) + '.xls'
    else:
        file_name = export_path + '.xls'

    # 保存xls文件
    workbook.save(file_name)


# 获取用户信息步骤（这里用户信息有两个部分组成；一部分是简单信息，一部分是详细信息）：
# 1.首先获取第一个用户的简单信息（open_id），以此为参数获取此用户的详细信息（creat_time）
# 2.获取剩余其它用户的open_id（需要第一个用户信息作参数）
# 3.获取剩余用户详细信息
# 4.构建用户信息字典，输出到Excel文件

# 示例步骤

# 1.登录
driver = longin_weixingzh()

# 2.解析当前网页，用BeautifulSoup解析
# Soup = parase_by_beautifulsoup(driver.page_source)
Soup = BeautifulSoup(driver.page_source, 'lxml')

# 3.获取当前网页信息（cookies/url/token）
html_headers_info = get_current_pageinfo(driver)
print(html_headers_info)

# 4.获取用户总人数（字符串）
fans_num = get_numbers_of_fans(Soup)
print(Soup)

# 5.获取第一个用户信息（需要open_id和createtime）
first_user_info = get_latest_user_info(html_headers_info['token'], html_headers_info['cookies'], Soup)[0]
print(first_user_info)
print('<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<')

# 6.获取剩余用户的简单信息（获得所有用户openid列表）
users_openid = get_users_simple_info(html_headers_info['url'], html_headers_info['token'], html_headers_info['cookies'],
                                     '50', first_user_info)
print('<<<<<users_openid<<<<<<<<<<<<<<<<<<')
print(users_openid)

# 7.获取用户详细信息
all_user_info_list = []

for user_info in users_openid:
    req = get_more_info(html_headers_info['token'], html_headers_info['cookies'], user_info)
    json_date = json.loads(req)
    user_data = json_date['user_list']['user_info_list'][0]
    all_user_info_list.append(user_data)

print('<<<<<<<<<all_user_info_list<<<<<<<<')
print(all_user_info_list)

# 8.将数据导入Excel
export_to_excel(all_user_info_list)
