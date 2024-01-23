import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

brows = webdriver.Chrome()
target_url="https://ethglobal.com/showcase"

index = 0
brows.get(target_url)
workbook = openpyxl.Workbook()

sheet = workbook.active

sheet['A1'] = 'Name'
sheet['B1'] = 'URL'
sheet['C1'] = 'Description'
sheet['E1']='Event'
sheet['I1']='Winner'

row_index = 2
page_counter = 0

try:
    while True:
        elements = brows.find_elements(By.XPATH, '//*[@class="block border-2 border-black rounded overflow-hidden relative"]')

        if index >= len(elements):
            try:
                next_page = brows.find_element(By.XPATH,'(//*[contains(@class,"w-10 h-10 border-2 flex items-center ")])[2]')
                next_page_url = next_page.get_attribute('href')
                if next_page_url is None:
                    break
                else:
                    brows.get(next_page_url)
                    index = 0
                    time.sleep(5)
                    continue

            except:
                break

        elements[index].click()
        time.sleep(3)

        parsed_names = brows.find_elements(By.XPATH, '//*[@class="text-4xl lg:text-5xl max-w-2xl mb-4"]')
        for name in parsed_names:
            sheet[f'A{row_index}'] = name.text
            row_index += 1

            parsed_url = brows.find_element(By.XPATH, '//*[contains(text(),"Source Code")]')
            href=parsed_url.get_attribute('href')
            if href =='https://github.com/':
                sheet[f'B{row_index}'] = 'URL does not exist'
            else:
                sheet[f'B{row_index}'] = href

        parsed_descriptions = brows.find_elements(By.XPATH,'/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[2]/div[1]/p[1]')
        for description in parsed_descriptions:
            sheet[f'C{row_index}'] = description.text
            row_index += 1

        parsed_events=brows.find_elements(By.XPATH,'//*[contains(@class,"inline-flex overflow")]')
        try:
            for event in parsed_events:
                sheet[f'E{row_index}']=event.text
                row_index += 1
        except:
            print('zalupaV2')

        parsed_winners=brows.find_elements(By.XPATH,'//*[@class="font-normal"]')
        if parsed_winners:
            for winner in parsed_winners:
                sheet[f'I{row_index}'] = winner.text
                row_index += 1
        else:
            sheet[f'I{row_index}'] = "Nothing in Winner"
            row_index += 1

        brows.back()
        time.sleep(5)
        index += 1

except:
    print('zalupa')

finally:
    workbook.save('parsed_data.xlsx')
    brows.close()
    brows.quit()