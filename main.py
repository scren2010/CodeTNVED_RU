import time
import xlsxwriter

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.remote.webdriver import WebDriver


def attach_to_session(executor_url, session_id):
    original_execute = WebDriver.execute

    def new_command_execute(self, command, params=None):
        if command == "newSession":
            # Mock the response
            return {'success': 0, 'value': None, 'sessionId': session_id}
        else:
            return original_execute(self, command, params)

    WebDriver.execute = new_command_execute
    driver = webdriver.Remote(command_executor=executor_url, desired_capabilities={})
    driver.session_id = session_id
    WebDriver.execute = original_execute
    return driver


def main():
    driver = webdriver.Chrome("D:/Downloads/chromedriver.exe")
    # driver = webdriver.Firefox()
    executor_url = driver.command_executor._url
    session_id = driver.session_id
    bro = attach_to_session(executor_url, session_id)
    bro.get('https://tnved.info/TnvedTree')

    workbook = xlsxwriter.Workbook('Товарная позиция (10 символа).xlsx')
    bold = workbook.add_format({'bold': True})
    worksheet = workbook.add_worksheet("Товарная позиция (4 символа)")
    worksheet.write('A1', 'Код', bold)
    worksheet.write('B1', 'Название', bold)
    worksheet.write('C1', 'Документ', bold)
    worksheet.write('D1', 'Подакцизный товар', bold)
    worksheet.write('E1', 'Тип товара', bold)
    worksheet.write('F1', 'Оформляется в УП', bold)
    worksheet.write('H1', 'Статус', bold)
    worksheet.write('I1', 'Раздел', bold)
    worksheet.write('J1', 'Группа', bold)
    worksheet.write('K1', 'Позиция', bold)
    worksheet.write('L1', 'Субпозиция', bold)
    worksheet.write('M1', 'Суб-субпозиция', bold)

    ul = driver.find_element_by_class_name("tree-item-block")
    razdel = ul.find_elements_by_class_name('tree-item')[20:21]

    t = 1
    time.sleep(3)
    nums = 2
    while razdel:
        s = []
        for i in razdel:
            code = i.find_element_by_class_name('tree-code')
            element = i.find_element_by_class_name('caret')
            driver.execute_script("arguments[0].click();", element)
            print(code.text)
            if len(code.text) == 10:
                div = i.find_element_by_class_name('tree-item-body')
                worksheet.write(f'A{nums}', code.text)
                worksheet.write(f'B{nums}', div.text)
                nums += 1
            time.sleep(t)
            try:
                ul = i.find_element_by_class_name("tree-item-block")
                li = ul.find_elements_by_class_name('tree-item')
                s += li
            except NoSuchElementException:
                pass
        razdel = [j for j in s]
    workbook.close()




if __name__ == "__main__":
    main()
