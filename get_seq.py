from selenium import webdriver
import time
import numpy
import xlrd
from multiprocessing.dummy import Pool
import xlwt
import configparser
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import hydra
from omegaconf import DictConfig, OmegaConf


def find(i,workbook,sheet1,sheet2,username,password):
    print(f'正在寻找第{i}个毒株的序列：{sheet1.cell(i, 1).value}')
    strain = sheet1.cell(i, 1).value
    sheet2.write(i, 1, sheet1.cell(i, 1).value)
    if '/' not in strain:
        return
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches",["enable-logging"])
    browser = webdriver.Chrome(options=options)
    browser.get("https://platform.epicov.org/epi3/frontend")
    tag = True
    #print('正在输入用户名')
    while tag:
        try:
            browser.find_element(by=By.XPATH, value=
            "/html/body/form/div[5]/div/div[2]/div/div/div/div[1]/div/div/div[2]/input[1]").send_keys(username)
            tag = False
        except:
            time.sleep(1)
    tag = True
    #print('正在输入密码')
    while tag:
        try:
            browser.find_element(by=By.XPATH, value=
            '/html/body/form/div[5]/div/div[2]/div/div/div/div[1]/div/div/div[2]/input[2]').send_keys(password)
            tag = False
        except:
            time.sleep(1)
    tag = True
    #print('正在点击登录')
    while tag:
        try:
            browser.find_element(by=By.XPATH, value=
            '/html/body/form/div[5]/div/div[2]/div/div/div/div[1]/div/div/div[2]/input[3]').click()
            tag = False
        except:
            time.sleep(1)
    tag = True

    # 输入accession并查找
    #print('正在输入strain')
    while tag:
        try:
            browser.find_element(by=By.XPATH, value=
            '/html/body/form/div[5]/div/div[2]/div/div[2]/div[2]/div[1]/table[3]/tbody/tr/td[2]/div/div[1]/input').send_keys(
                sheet1.cell(i, 1).value)
            tag = False
        except:
            time.sleep(1)
    tag = True
    #print('正在点击HA')
    while tag:
        try:
            ha_button = browser.find_element(by=By.XPATH, value=
            '/html/body/form/div[5]/div/div[2]/div/div[2]/div[2]/div[9]/table[3]/tbody/tr/td[2]/div/div[1]/div[4]/input')
            browser.execute_script("arguments[0].click();", ha_button)
            tag = False
        except:
            time.sleep(1)
    tag = True
    time.sleep(5)
    #print('正在点击查找')
    while tag:
        try:
            button_search_list = browser.find_elements(by=By.CLASS_NAME, value='sys-form-button')
            button_search = button_search_list[2]
            browser.execute_script("arguments[0].click();", button_search)
            tag = False
        except:
            time.sleep(1)
    tag = True
    xx1 = 0
    while tag:
        try:
            list_a = browser.find_elements(by=By.XPATH, value=
            '/html/body/form/div[5]/div/div[2]/div/div[2]/div/div[1]/div[3]/table/tbody[2]/tr')
            max_len_index = 0
            max_len = 0
            for j in range(0, len(list_a)):
                tx = browser.find_element(by=By.XPATH,value=f'/html/body/form/div[5]/div/div[2]/div/div[2]/div/div[1]/div[3]/table/tbody[2]/tr[{j+1}]/td[3]').text
                if strain== tx:
                    ll = list_a[j].find_elements(by=By.XPATH, value='./td')
                    if max_len < int(ll[9].text.replace(',', '')):
                        max_len = max(max_len, int(ll[9].text.replace(',', '')))
                        max_len_index = j
            browser.find_element(by=By.XPATH, value=
            f'/html/body/form/div[5]/div/div[2]/div/div[2]/div/div[1]/div[3]/table/tbody[2]/tr[{max_len_index+1}]').click()
            tag = False
        except:
            if xx1>120:
                return
            xx1 = xx1+1
            time.sleep(1)
    tag = True
    xx1 = 0
    #print('正在寻找最优答案')
    while tag:
        try:
            name = browser.find_elements(by=By.TAG_NAME, value='iframe')
            browser.switch_to.frame(name[0])
            tr_list = browser.find_elements(by=By.XPATH, value=
            '/html/body/form/div[5]/div/div[2]/div/div[2]/div/div[10]/table/tbody/tr')
            tag = False
        except:
            time.sleep(1)
    tag = True
    first = True
    col = 6
    length = 0
    answer = ''
    for tr in tr_list:
        if first:
            first = False
            continue
        strip = tr.text.split('\n')
        if strip[0] == 'HA':
            if len(strip) >= 5:
                if int(strip[2]) > length:
                    length = int(strip[2])
                    answer = strip[4]
                    # #print(f'accession是{answer}')
            else:
                if int(strip[2]) > length:
                    length = int(strip[2])
                    answer = strip[3]
    #print(f'最长的accession是{answer}，长度为{length}')
    if len(strip) >= 5:
        browser.get(f'https://www.ncbi.nlm.nih.gov/nuccore/{answer}?report=GenBank')
        tag = True
        while tag:
            try:
                st = browser.find_element(by=By.ID, value=
                f'feature_{answer}.1_CDS_0').text.split(
                    'translation="')[1].replace('\n', '').replace("'", "").replace(' ', '')
                #print(st)
                tag = False
            except:
                time.sleep(1)
        tag = True
        sheet2.write(i,2,'NCBI')
    else:
        for j in range(1, len(tr_list)):
            if str(length) in tr_list[j].text and 'HA' in tr_list[j].text:
                handle = browser.current_window_handle
                while tag:
                    try:
                        link_list = browser.find_elements(by=By.CLASS_NAME, value='sys-form-fi-link')
                        if len(link_list) == 2 * (len(tr_list) - 1):
                            link = link_list[2 * j - 1]
                        else:
                            link = link_list[j - 1]
                        browser.execute_script("arguments[0].click();", link)
                        tag = False
                    except:
                        time.sleep(1)
                tag = True
                #print('正在切换至跳转的页面')
                while tag:
                    try:
                        name = browser.find_elements(by=By.TAG_NAME, value='iframe')
                        browser.switch_to.frame(name[0])
                        s = browser.find_element(by=By.XPATH, value=
                        '/html/body/form/div[5]/div/div[2]/div/div[2]/div/table/tbody/tr[5]/td[1]/pre').text
                        check_text = browser.find_element(by=By.XPATH,
                                                          value='/html/body/form/div[5]/div/div[2]/div/div[2]/div/table/tbody/tr[2]').text.replace(
                            '\n', '')
                        tag = False
                    except:
                        time.sleep(1)
                tag = True
                split = s.split()
                string1 = ''
                for k in range(0, len(split)):
                    string1 = (string1 + split[k])
                st = ''.join([i for i in string1 if not i.isdigit()])
                #print(st)
                sheet2.write(i,2,'GISAID')
                # #print('answer:', answer)
                # #print('check_text:', check_text, '\n')
                break
    sheet2.write(i,3,st)
    workbook.save('tt3.xls')
    browser.quit()

@hydra.main(version_base=None, config_name="config", config_path="./config/")
def main(cfg: DictConfig):
    xls = xlrd.open_workbook_xls('ans.xls')

    # print(f'打开了{xls}文件')
    names = xls.sheet_names()
    sheet1 = xls.sheet_by_name(names[0])
    nrows = sheet1.nrows
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet2 = workbook.add_sheet('test1')

    # seq_list = numpy.linspace(1, nrows - 1, nrows-1, dtype=int)
    # seq_list = numpy.linspace(0, nrows - 1, nrows, dtype=int)
    # tolist = seq_list.tolist()
    for i in range(1, nrows):
        find(i,workbook,sheet1,sheet2,cfg.username,cfg.password)
    # pool = Pool(1)
    # pool.map(find, tolist)
    # pool.close()
    # pool.join
    workbook.save('result.xls')

if __name__ == '__main__':
    main()




