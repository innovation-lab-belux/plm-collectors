import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import glob
from collections import Counter
import clipboard
import json

appended_data = []

os.chdir('Recipes')
FileList = glob.glob('*.xlsx')

specs=[]
iterances=[]

r=0
for File in FileList:
    r+=1
    
    df = pd.read_excel(File)
    identifier = df['Specification'][0]
    recipe = '{ "id":"'+identifier+'","ingredients":['

    for i in range(len(df['Specification'])):

        spec = df['Specification'][i]
        brix = str(df['Brix, uncorrected'][i])
        acid = str(df['Acid, CA pH 8,1'][i])
        quantity = str(df['Logistic Quantity'][i])

        ingredient=''
        try:
            with open('data_'+spec+'.json', 'r') as file:
                data = json.load(file)
                ingredient = json.dumps(data)
                ingredient = ingredient.replace("caillr@nyp6.onmicrosoft.com","no information")

        except:
            n=0

        if ingredient!='':
            ingredient = ingredient[:-1]+',"quantity (kg)": "'+quantity+'", "acidity (ph)":"'+acid+'", "sweetness (brix)": "'+brix+'"},'
            recipe = recipe+ingredient


    recipe=recipe[:-1]
    recipe=recipe+"]}"
    print(r)

    y = json.loads(recipe)
    with open('dataFinal_'+identifier+'.json', 'w') as f:
        json.dump(y, f)

    print(r)
        

'''
print(Counter(iterances))
import pandas as pd
df = pd.DataFrame.from_dict(Counter(iterances), orient='index').reset_index()
df.columns =['column1', 'colum2']
df = df.sort_values(by=['column1'],ascending=False)
print(df)
df.to_csv("freq.csv", encoding='utf-8', index=False)
'''

driver = webdriver.Chrome()

def visit_initial_url():
    driver.get("https://www.example.com")
    driver.execute_script("window.open('https://uxm-sap-build-work-zone-adv-oxa2dgkx.workzone.cfapps.eu10.hana.ondemand.com/site#workzone-notification?sap-app-origin-hint=&/groups/ZOz7wRSBggz6pPTIvuI000/workpage_tabs/WHiRAT3Z7r0xQyKhfVS000', '_blank');")

def extract_excel(spec):
    time.sleep(1)

    actions = ActionChains(driver)
    element = driver.find_element(By.TAG_NAME,'body') 

    #scroll
    x = 201-756
    y = 531-403
    actions.move_to_element_with_offset(element, x, y)
    actions.click_and_hold().move_by_offset(0, -466).pause(0.1).release().perform()
    time.sleep(1)

    #grouping
    x = 74 - 756
    y = 238 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.click().perform()
    time.sleep(1)

    #substance type
    x = 153 - 756
    y = 536 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.double_click().perform()
    time.sleep(0.1)
    actions.click().perform()
    time.sleep(1)

    #copy type
    x = 770 - 756
    y = 390 - 403 
    actions.move_to_element_with_offset(element, x, y)
    actions.context_click().pause(1).send_keys(Keys.RETURN).pause(1).perform()
    time.sleep(0.1)
    type = clipboard.paste()

    #move to product
    x = 156 - 756
    y = 565 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.double_click().perform()
    time.sleep(0.1)
    actions.click().perform()
    time.sleep(1)

    #copy product
    x = 605 - 756
    y = 390 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.context_click().pause(1).send_keys(Keys.RETURN).pause(1).perform()
    time.sleep(0.1)
    product = clipboard.paste()

    # close product
    x = 75 - 756
    y = 238 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.click().perform()
    time.sleep(1)

    #open aroma
    x = 74 - 756
    y = 313 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.click().perform()
    time.sleep(1)

    #open taste and aroma
    x = 172 - 756
    y = 336 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.double_click().perform()
    time.sleep(0.1)
    actions.click().perform()
    time.sleep(1)

    #copy taste
    x = 746 - 756
    y = 392 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.context_click().pause(1).send_keys(Keys.RETURN).pause(1).perform()
    time.sleep(0.1)
    aroma = clipboard.paste()

    #close aroma
    x = 75 - 756
    y = 314 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.click().perform()
    time.sleep(1)

    #open colour
    x = 75 - 756
    y = 337 - 403
    actions.move_to_element_with_offset(element, x, y)

    actions.click().perform()
    time.sleep(1)

    #open colour detail
    x = 162 - 756
    y = 364 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.double_click().perform()
    time.sleep(0.1)
    actions.click().perform()
    time.sleep(1)

    #copy colour
    x = 841 - 756
    y = 392 - 403
    actions.move_to_element_with_offset(element, x, y)
    actions.context_click().pause(1).send_keys(Keys.RETURN).pause(1).perform()
    colour = clipboard.paste()

    #save json
    x =  '{ "spec":"'+spec+'","substance_type":"'+type+'","product_type":"'+product+'","colour":"'+colour+'","aroma":"'+aroma+'"}'
    y = json.loads(x)
    with open('data_'+spec+'.json', 'w') as f:
        json.dump(y, f)

    time.sleep(1) 
    
    # caillr@nyp6.onmicrosoft.com 

visit_initial_url()
input()
time.sleep(1)

def move_to_coordinates(x,y,actions,driver):
    x_offset = x-756
    y_offset = y-430
    element = driver.find_element(By.TAG_NAME,'body') 
    actions.move_to_element_with_offset(element, x_offset, y_offset)
    return actions


'''
for i in range(len(specs)):

    spec = specs[i]
    url = "https://vlcspcomdoe.devsys.net.sap:44311/sap/bc/gui/sap/its/webgui?~transaction=*ZRM_ASS_SUBID_DISPL%20P_SUBID="+spec+";DYNP_OKCODE=SEARCH2#"
    print(url)
    print(i)

    if i>1640:
        print("")

        time.sleep(0.5)
        driver.execute_script("window.open('"+url+"', '_blank')")
        driver.switch_to.window(driver.window_handles[-1])
        extract_excel(spec)
        time.sleep(0.5)
        driver.close()
        driver.switch_to.window(driver.window_handles[-1])


    # caillr@nyp6.onmicrosoft.com 
'''