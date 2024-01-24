import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
from openpyxl.styles import Alignment

brows = webdriver.Chrome()
target_url = "https://ethglobal.com/showcase"

brows.get(target_url)
workbook = openpyxl.Workbook()

sheet = workbook.active

headers = ['Name', 'URL', 'Description', 'Event', 'Winner']
sheet.append(headers)

column_widths = {'A': 30, 'B': 20, 'C': 50, 'D': 20, 'E': 20, 'I': 20}
for col, width in column_widths.items():
    sheet.column_dimensions[col].width = width

for row in sheet.iter_rows():
    for cell in row:
        cell.alignment = Alignment(wrap_text=True)

try:
    index = 0
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

        project_row_index = sheet.max_row + 1

        elements[index].click()
        time.sleep(3)

        name = brows.find_element(By.XPATH, '//*[@class="text-4xl lg:text-5xl max-w-2xl mb-4"]').text
        sheet[f'A{project_row_index}'] = name

        try:
            parsed_url = brows.find_element(By.XPATH, '//*[contains(text(),"Source Code")]')
            href = parsed_url.get_attribute('href')
            url = href if href != 'https://github.com/' else 'URL does not exist'
        except:
            url = 'URL not found'
        sheet[f'B{project_row_index}'] = url

        description = brows.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div[2]/div[1]/p[1]').text
        sheet[f'C{project_row_index}'] = description

        event = brows.find_element(By.XPATH, '//*[contains(@class,"inline-flex overflow")]').text
        sheet[f'D{project_row_index}'] = event

        parsed_winners = brows.find_elements(By.XPATH, '//*[@class="font-normal"]')
        if parsed_winners:
            winners_text = ""
            for winner in parsed_winners:
                winners_text += winner.text + " "
            winners_text = winners_text.strip()
            sheet[f'I{project_row_index}'] = winners_text
        else:
            sheet[f'I{project_row_index}'] = "Nothing in Winner"

        brows.back()
        time.sleep(5)
        index += 1

except :
    print('zalupa')

finally:
    workbook.save('parsed_data.xlsx')
    brows.close()
    brows.quit()
