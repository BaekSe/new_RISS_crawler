import os, time, glob, shutil
from selenium import webdriver
from math import ceil
import re

def load_journal_over(driver, page, pageScale, keyword, StartCount):

    url = 'http://www.riss.kr/search/Search.do?detailSearch=false&viewYn=OP&query='
    url2 = '&iStartCount='
    url3 = '&iGroupView=5&icate=re_a_over&colName=re_a_over&exQuery=&pageScale=500&strSort=RANK&order=%2FDESC&onHanja=false&keywordOption=0&searchGubun=true&p_year1=&p_year2=&dorg_storage=&mat_type=&mat_subtype=&fulltext_kind=&t_gubun=&learning_type=&language_code=&ccl_code=&language=&inside_outside=&fric_yn=&image_yn=&regnm=&gubun=&kdc=&resultSearch=false&listFlag=&h_groupByField='
    full_url = url + keyword + url2 + str(StartCount + page * pageScale) + url3
    driver.get(full_url)

    while '(검색결과 : 0 건)' ==  driver.find_element_by_xpath('//*[@id="level4_mainContent"]/form/div[1]/div/div/ul/li/span[2]').text:
        driver.refresh()
        driver.implicitly_wait(3)

    window_before = driver.window_handles[0]
    driver.find_element_by_xpath('//*[@id="allchk2"]').click()
    driver.find_element_by_xpath('//*[@id="level4_mainContent"]/form/div[3]/div[2]/div/div[1]/p[1]/span/a[1]').click()

    window_after = driver.window_handles[-1]
    driver.switch_to_window(window_after)
    driver.implicitly_wait(10)

    try:
        driver.find_element_by_xpath('//*[@id="radio-3"]').click()
    except:
        return

    driver.find_element_by_xpath('//*[@id="radio-8"]').click()
    driver.find_element_by_xpath('/html/body/form/div/div[2]/div[5]/a[1]/img').click()
    time.sleep(20)
    
    driver.close()
    driver.switch_to_window(window_before)
    return


if __name__ == '__main__':
    f = open('keyword.txt', 'r')
    keyword = f.readline()
    f.close()
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
    })
    driver = webdriver.Chrome('./chromedriver.exe', chrome_options=options)
    driver.implicitly_wait(3)
    StartCount = int()

    if not os.path.exists('../journal_over'):
        os.mkdir('../journal_over')

    try:
        F = open('journal_over_startcount.txt', 'r')
        StartCount = int(F.readline())
    except:
        StartCount = 0
    pageScale = 500

    driver.get('http://www.riss.kr/search/Search.do?detailSearch=false&viewYn=OP&query='+keyword+'&iStartCount='+str(StartCount) + '&iGroupView=5&icate=re_a_over&colName=re_a_over&exQuery=&pageScale=500&strSort=RANK&order=%2FDESC&onHanja=false&keywordOption=0&searchGubun=true&p_year1=&p_year2=&dorg_storage=&mat_type=&mat_subtype=&fulltext_kind=&t_gubun=&learning_type=&language_code=&ccl_code=&language=&inside_outside=&fric_yn=&image_yn=&regnm=&gubun=&kdc=&resultSearch=false&listFlag=&h_groupByField=')
    results = driver.find_element_by_xpath('/html/body/div[1]/div[4]/div[2]/div[3]/form/div[1]/div/div/ul/li/span[2]').text
    
    results = int(''.join(re.findall('\\d', results))) - StartCount
    print('검색 결과:', results, '+', StartCount, '건')

    for i in range(ceil(results / pageScale)):
        print(str(i * pageScale + StartCount) + "개 다운 완료")
        with open('journal_over_startcount.txt', 'w') as F:
            F.write(str(StartCount + i * pageScale))
        load_journal_over(driver, i, pageScale, keyword, StartCount)
        for filename in os.listdir(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')):
            if filename.startswith("myCabinet") and ".crdownload" not in filename:
                try:
                    os.rename(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads', filename), os.path.join("../journal_over", "journal_over_data"+str(StartCount + i * pageScale) + filename[-4:]))
                except:
                    os.rename(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads', filename), os.path.join("../journal_over", "journal_over_data"+str(StartCount + i * pageScale + 1) + filename[-4:]))
    driver.quit()
