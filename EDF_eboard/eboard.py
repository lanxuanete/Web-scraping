# coding=utf-8
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
import time
from browsermobproxy import Server
import random
from selenium.webdriver.common.by import By
import json
import pandas as pd

def get_ua():
    first_num = random.randint(55, 76)
    third_num = random.randint(0, 3800)
    fourth_num = random.randint(0, 140)
    os_type = [
        '(iPhone; CPU iPhone OS 13_2_3 like Mac OS X)', '(iPad; CPU OS 11_0 like Mac OS X)',
        '(Linux; Android 6.0.1; Moto G (4))'
    ]
    chrome_version = 'Chrome/{}.0.{}.{}'.format(first_num, third_num, fourth_num)
    ua = ' '.join(['Mozilla/5.0', random.choice(os_type), 'AppleWebKit/537.36',
                   '(KHTML, like Gecko)', chrome_version, 'Safari/537.36']
                  )
    return ua


server = Server(path="D:\\PycharmProjects\\browsermob-proxy-2.1.4\\bin\\browsermob-proxy.bat")
server.start()
proxy = server.create_proxy()
print(proxy.proxy)

chromedrive_path = r"D:\chromedriver.exe"
binary_location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
options = Options()
options.binary_location = binary_location
options.add_experimental_option('excludeSwitches', ['enable-automation'])
options.add_argument('--incognito')
options.add_argument('disable-infobars')
options.add_argument('log-level=3')
options.add_argument("--auto-open-devtools-for-tabs")
options.add_argument('--proxy-server={0}'.format(proxy.proxy))
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-urlfetcher-cert-requests')
# options.add_argument('--headless')
user_agent = get_ua()
options.add_argument('user-agent=%s' % (user_agent))


driver = Chrome(chrome_options=options, executable_path=chromedrive_path)
proxy.new_har("huaruntong", options={'captureHeaders': True, 'captureContent': True})
driver.get("https://eboard.edf.fr/entreprises/index.html#/external/marches/details?marcheIds=585610&marcheIds=585652&indiceSelectionChanged=true")

driver.maximize_window()
time.sleep(5)

##login
driver.find_element(By.XPATH,'//*[@id="tc_privacy_button"]').click()
time.sleep(5)
driver.find_element(By.XPATH,'//*[@id="idToken1"]').send_keys('broy@ressource-consulting.fr')
print("send username")
time.sleep(1)
driver.find_element(By.XPATH,'//*[@id="idToken2"]').send_keys('RS5Saulnier!')
print("send Password")
time.sleep(1)
driver.find_element(By.XPATH,'//*[@id="loginButton_0"]').click()
print("login")
time.sleep(15)
driver.find_element(By.XPATH,'/html/body/app-root/div[3]/div/div/div[2]/div[1]/button').click()
time.sleep(5)
driver.find_element(By.XPATH,'//*[@id="main-navbar-icons-left"]/li[6]').click()
time.sleep(10)
# ----------------------
#
result = proxy.har

for entry in result['log']['entries']:
    _url = entry['request']['url']
    if "https://eboard.edf.fr/entreprises/api/getMarchesDetailData?ids=584425&ids=584419&ids=585010&ids=585011&ids=585017&ids=584431&ids=585023&ids=585015&ids=585016&ids=585018&ids=585012&ids=585013&ids=585014&ids=585021&ids=585019&ids=585024&ids=585022&ids=585020&ids=585025&ids=535469&ids=563677&ids=508347&ids=563714&ids=535469&ids=535469&ids=535469&ids=535469&ids=535469&ids=535469&ids=535469&ids=535469&ids=535469&ids=535469" in _url:
        _response = entry['response']
        _content = _response['content']['text']
        # print("---------请求响应内容-----------：",_content)
        ## type( _content)：<class 'dict'>
        ## type(_content['text'])： <class 'str'>

server.stop()
driver.quit()

eex_dict=json.loads(_content)

i=0
for eex_dataGraphs_Gaz in eex_dict['dataGraphs']['Gaz']:
    df_dataGraphs_Gaz=pd.DataFrame.from_records(eex_dataGraphs_Gaz)
    print(df_dataGraphs_Gaz)
    lable_name=eex_dict['dataGraphs']['Gaz'][i][0]['shortLabel']
    df_dataGraphs_Gaz.to_csv(r'D:\PycharmProjects\EDF\result\process_file\eboard_gaz.csv', mode='a', header=True)
    print(i,lable_name,"a fini")
    i+=1
