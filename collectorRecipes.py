import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
df = pd.read_excel('Recipes.xlsx')
driver = webdriver.Chrome()

def visit_initial_url():
    driver.get("https://www.example.com")
    driver.execute_script("window.open('https://uxm-sap-build-work-zone-adv-oxa2dgkx.workzone.cfapps.eu10.hana.ondemand.com/site#workzone-notification?sap-app-origin-hint=&/groups/ZOz7wRSBggz6pPTIvuI000/workpage_tabs/WHiRAT3Z7r0xQyKhfVS000', '_blank');")

def extract_excel(url):
    driver.get(url)
    time.sleep(5)

    #input_field = driver.find_element(By.ID,'tree#C115#4#1#1#i-text') 
    
    '''actionsClick = ActionChains(driver)
    actionsClick.move_to_element()
    actionsClick.double_click().perform()
    
    input_field.click().perform()
    time.sleep(0.1)
    input_field.click().perform()
    time.sleep(6)
'''
    x_offset = 399-756
    y_offset = 319-400
    actions = ActionChains(driver)
    element = driver.find_element(By.TAG_NAME,'body') 

    actions.move_to_element_with_offset(element, x_offset, y_offset)
    actions.click().perform()
    time.sleep(0.5)
    actions.send_keys(Keys.RETURN).perform()
    time.sleep(1) 

    x_offset = 411-756
    y_offset = 287-400
    actions = ActionChains(driver)
    element = driver.find_element(By.TAG_NAME,'body') 

    actions.move_to_element_with_offset(element, x_offset, y_offset)
    actions.click().perform()
    time.sleep(0.5)
    actions.send_keys(Keys.RETURN).perform()
    time.sleep(1)


visit_initial_url()
input()

print(df.size)
for i in range(len(df)):
    if i>0:
        url = df['url'][i]
        print(url)
        print(i)
        print("")
        time.sleep(2)
        driver.execute_script("window.open('"+str(url)+"', '_blank')")
        driver.switch_to.window(driver.window_handles[-1])
        extract_excel(url)
        time.sleep(0.5)
        driver.close()
        driver.switch_to.window(driver.window_handles[-1])

    #27
