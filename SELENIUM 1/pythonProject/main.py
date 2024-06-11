import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def WaitById(driver, nameId):
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, nameId))
        )
        return driver.find_element(By.ID, nameId)
    except:
        driver.quit()
        exit()


def WaitByClass(driver, nameClass):
    try:
        wait = driver.find_elements(By.CLASS_NAME, nameClass)
        return wait
    except:
        driver.quit()
        exit()

def WaitByClassSingle(driver, nameClass):
    try:
        wait = WebDriverWait(driver, 10)
        wait = wait.until(EC.presence_of_element_located((By.CLASS_NAME, nameClass)))
        return wait
    except:
        driver.quit()
        exit()

def WaitByCSS(driver, nameCSS):
    try:
        wait = driver.find_element(By.CSS_SELECTOR, nameCSS)
        return wait
    except:
        driver.quit()
        exit()


driver = webdriver.Chrome()
i = 0
curPage = 1
WebLink = "https://www.tokopedia.com/search?navsource=thematic&ob=5&page=" + str(
    curPage) + "&q=sambel&source=search&srp_component_id=04.06.00.00&srp_disco_url=beli-lokal&srp_page_id=21239&srp_page_title=Beli%20Lokal"
driver.get(WebLink)
driver.implicitly_wait(10)
asa = 0
aced = 100
file = open("link.txt", "a")
list_web = []
list_toko = []
while i <= 1000:
    time.sleep(1)
    while asa < 5000:
        driver.execute_script("window.scrollTo("+str(asa)+", "+str(aced)+");")
        asa = asa+100
        aced = aced + 100
        time.sleep(0.1)
    time.sleep(1)
    search = WaitByClass(driver, "css-19oqosi")
    for link in search:
        temp = (link.find_element(By.TAG_NAME, "a")).get_attribute("href")
        list_web.append((link.find_element(By.TAG_NAME, "a")).get_attribute("href"))
        i = i + 1
    print(i)
    time.sleep(1)
    curPage = curPage + 1
    WebLink = "https://www.tokopedia.com/search?navsource=thematic&ob=5&page=" + str(
        curPage) + ("&q=sambel&source=search&srp_component_id=04.06.00.00&srp_disco_url=beli-lokal&srp_page_id=21239"
                    "&srp_page_title=Beli%20Lokal")
    driver.get(WebLink)
    asa = 0
    aced = 100
    driver.implicitly_wait(10)
i = 0
for link in list_web:
    i = i+1
    print(i)
    driver.get(link)
    driver.implicitly_wait(20)
    driver.execute_script("window.scrollTo(0, 500);")
    search = WaitByClassSingle(driver, "css-1sl4zpk")
    search = search.get_attribute("href")
    if search not in list_toko:
        file.write(search + "\n")
        list_toko.append(search)
file.close()
driver.quit()
