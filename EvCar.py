import requests as req
import json
import os
import re
import xlwt
from selenium import webdriver


'''
实现参考：https://www.cnblogs.com/kangz/p/10011348.html
'''
def get_brand_list():
    print("get brand list...")
    url = "https://ev.autohome.com.cn/ashx/AjaxIndexCarFind.ashx?type=1"
    resp = req.get(url)
    ret = json.loads(resp.content)['result']['branditems']
    return ret


'''
获取所有车型的网页，保存到本地
'''
def get_series():
    print("get series list...")
    brands = get_brand_list()
    # fixme: 测试，指定爬取的品牌字母范围
    allow_leeters = [chr(i) for i in range(ord("A"), ord("B") + 1)]
    for b in brands:
        url = "https://ev.autohome.com.cn/ashx/AjaxIndexCarFind.ashx?type=3&value={}".format(b['id'])
        resp = req.get(url)
        ret = json.loads(resp.content)['result']['factoryitems']
        if b['bfirstletter'] in allow_leeters:
            for series in ret:
                for model in series['seriesitems']:
                    name = "{}.html".format(model['id'])
                    # 车系列页面
                    series_url = "https://car.autohome.com.cn/config/series/{}".format(name)
                    print("{} --> {}".format(model['name'], series_url))
                    # print(req.get(series_url).content)
                    resp = req.get(series_url)
                    file = open("d:\\autoHome\\html\\{}".format(model['id']), "a", encoding="utf-8")
                    file.write(str(resp.content, "utf-8"))
                    file.close()


'''
解析出每个车型的关键js并拼装成一个html,保存到本地
'''
def get_js():
    print("get js...")
    root_path = "D:\\autoHome\\html\\"
    files = os.listdir(root_path)
    for file in files:
        print("fileName==" + file.title())
        text = ""
        for fi in open(root_path + file, 'r', encoding="utf-8"):
            text = text + fi
        else:
            print("fileName==" + file.title())
        # 解析数据的json
        alljs = ("var rules = '2';"
                 "var document = {};"
                 "function getRules(){return rules}"
                 "document.createElement = function() {"
                 "      return {"
                 "              sheet: {"
                 "                      insertRule: function(rule, i) {"
                 "                              if (rules.length == 0) {"
                 "                                      rules = rule;"
                 "                              } else {"
                 "                                      rules = rules + '#' + rule;"
                 "                              }"
                 "                      }"
                 "              }"
                 "      }"
                 "};"
                 "document.querySelectorAll = function() {"
                 "      return {};"
                 "};"
                 "document.head = {};"
                 "document.head.appendChild = function() {};"

                 "var window = {};"
                 "window.decodeURIComponent = decodeURIComponent;")
        try:
            js = re.findall('(\(function\([a-zA-Z]{2}.*?_\).*?\(document\);)', text)
            for item in js:
                alljs = alljs + item
        except Exception as e:
            print('makejs function exception')

        newHtml = "<html><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /><head></head><body>    <script type='text/javascript'>"
        alljs = newHtml + alljs + " document.write(rules)</script></body></html>"
        f = open("D:\\autoHome\\newhtml\\" + file + ".html", "a", encoding="utf-8")
        f.write(alljs)
        f.close()


'''
.解析出每个车型的数据json，比如var config  ,var option , var bag等
'''
def get_json():
    print("get json...")
    print("Start...")
    rootPath = "D:\\autoHome\\html\\"
    files = os.listdir(rootPath)
    for file in files:
        print("fileName==" + file.title())
        text = ""
        for fi in open(rootPath + file, 'r', encoding="utf-8"):
            text = text + fi
        else:
            print("fileName==" + file.title())
        # 解析数据的json
        jsonData = ""
        config = re.search('var config = (.*?){1,};', text)
        if config != None:
            print(config.group(0))
            jsonData = jsonData + config.group(0)
        option = re.search('var option = (.*?)};', text)
        if option != None:
            print(option.group(0))
            jsonData = jsonData + option.group(0)
        bag = re.search('var bag = (.*?);', text)
        if bag != None:
            print(bag.group(0))
            jsonData = jsonData + bag.group(0)
        f = open("D:\\autoHome\\json\\" + file, "a", encoding="utf-8")
        f.write(jsonData)
        f.close()

'''
生成样式文件，保存到本地。这一步很重要，网站经过混淆处理，部分显示的文字是通过css样式拼出来的
chromedriver获取参考：https://blog.csdn.net/sdzhr/article/details/82660860
使用参考：https://www.cnblogs.com/hellosecretgarden/p/9206648.html
'''
def get_content():
    print("get content...")
    chrome = webdriver.Chrome("D:\Program Files\Chrome\chromedriver.exe")

    lists = os.listdir("D:/autoHome/newHtml/")
    for fil in lists:
        file = os.path.exists("D:\\autoHome\\content\\" + fil)
        if file:
            print('文件已经解析。。。' + str(file))
            continue
        print(fil)
        chrome.get("file:///D:/autoHome/newHtml/" + fil + "")
        text = chrome.find_element_by_tag_name('body')
        print(text.text)
        f = open("D:\\autoHome\\content\\" + fil, "a", encoding="utf-8")
        f.write(text.text)
        f.close()

    chrome.close()


'''
读取样式文件，匹配数据文件，生成正常数据文件
'''
def get_config():
    print("get config...")
    rootPath = "D:\\autoHome\\json\\"
    listdir = os.listdir(rootPath)
    for json_s in listdir:
        print(json_s.title())
        jso = ""
        # 读取json数据文件
        for fi in open(rootPath + json_s, 'r', encoding="utf-8"):
            jso = jso + fi
        content = ""
        # 读取样式文件
        spansPath = "D:\\autoHome\\content\\" + json_s.title() + ".html"
        for spans in open(spansPath, "r", encoding="utf-8"):
            content = content + spans
        print(content)
        # 获取所有span对象
        jsos = re.findall("<span(.*?)></span>", jso)
        num = 0
        for js in jsos:
            print("匹配到的span=>>" + js)
            num = num + 1
            # 获取class属性值
            sea = re.search("'(.*?)'", js)
            print("匹配到的class==>" + sea.group(1))
            spanContent = str(sea.group(1)) + "::before { content:(.*?)}"
            # 匹配样式值
            spanContentRe = re.search(spanContent, content)
            if spanContentRe != None:
                if sea.group(1) != None:
                    print("匹配到的样式值=" + spanContentRe.group(1))
                    jso = jso.replace(str("<span class='" + sea.group(1) + "'></span>"),
                                      re.search("\"(.*?)\"", spanContentRe.group(1)).group(1))
        print(jso)
        fi = open("D:\\autoHome\\newJson\\" + json_s.title(), "a", encoding="utf-8")
        fi.write(jso)
        fi.close()


targe_config = {"车型名称", "工信部纯电续航里程(km)", "实测续航里程(km)", "电池类型", "电池能量(kWh)",
                        "百公里耗电量(kWh/100km)", "快充时间(小时)", "快充电量百分比"}

'''
指定需要的字段，从数据文件中提取出来保存到excel中
'''
def save_2_xls():
    print("save to excel ...")
    rootPath = "D:\\autoHome\\newJson\\"
    workbook = xlwt.Workbook(encoding='ascii')  # 创建一个文件
    worksheet = workbook.add_sheet('汽车之家')  # 创建一个表
    files = os.listdir(rootPath)
    startRow = 0
    isFlag = True  # 默认记录表头
    for file in files:
        carItem = {}
        print("fileName==" + file.title())
        text = ""
        for fi in open(rootPath + file, 'r', encoding="utf-8"):
            text = text + fi
        # 解析基本参数配置参数，颜色三种参数，其他参数
        config = "var config = (.*?);"
        option = "var option = (.*?);var"
        bag = "var bag = (.*?);"

        configRe = re.findall(config, text)
        optionRe = re.findall(option, text)
        bagRe = re.findall(bag, text)
        for a in configRe:
            config = a
        print("++++++++++++++++++++++\n")
        for b in optionRe:
            option = b
            print("---------------------\n")
        for c in bagRe:
            bag = c

        motorItems = []
        try:
            config = json.loads(config)

            configItem = config['result']['paramtypeitems'][0]['paramitems']
            paramTypes = config['result']['paramtypeitems']
            for paramType in paramTypes:
                if paramType['name'] == "电动机":
                    motorItems = paramType['paramitems']
        except Exception as e:
            print(e)
            f = open("D:\\autoHome\\异常数据\\exception.txt", "a", encoding="utf-8")
            f.write(file.title() + "\n")
            continue

        # 解析基本参数
        for car in configItem:
            carItem[car['name']] = []
            for ca in car['valueitems']:
                if car['name'] in targe_config:
                    carItem[car['name']].append(ca['value'])
        for car in motorItems:
            carItem[car['name']] = []
            for ca in car['valueitems']:
                if car['name'] in targe_config:
                    carItem[car['name']].append(ca['value'])

        if isFlag:
            co1s = 0

            for co in carItem:
                if co in targe_config:
                    co1s = co1s + 1
                    worksheet.write(startRow, co1s, co)
            else:
                startRow = startRow + 1
                isFlag = False

        # 计算起止行号
        endRowNum = startRow + len(carItem['车型名称'])  # 车辆款式记录数
        for row in range(startRow, endRowNum):
            print(row)
            colNum = 0
            for col in carItem:
                if col in targe_config:
                    colNum = colNum + 1
                    print(str(carItem[col][row - startRow]), end='|')
                    worksheet.write(row, colNum, str(carItem[col][row - startRow]))

        else:
            startRow = endRowNum
    workbook.save('d:\\autoHome\\Mybook.xls')


def del_file(path):
    print("clear dir: {}".format(path))
    for i in os.listdir(path):
        path_file = os.path.join(path, i) # 取文件绝对路径
        if os.path.isfile(path_file):
            os.remove(path_file)
        else:
            del_file(path_file)


def prepare_dir(path):
    if os.path.isdir(path):
        del_file(path)
    else:
        print("make dir: {}".format(path))
        os.mkdir(path)


if __name__ == "__main__":
    prepare_dir("D:\\autoHome\\content")
    prepare_dir("D:\\autoHome\\html")
    prepare_dir("D:\\autoHome\\json")
    prepare_dir("D:\\autoHome\\newhtml")
    prepare_dir("D:\\autoHome\\newJson")
    prepare_dir("D:\\autoHome\\异常数据")

    get_series()
    get_js()
    get_json()
    get_content()
    get_config()
    save_2_xls()
