import time
from selenium import webdriver
def main():
    options = webdriver.ChromeOptions()
    # Path to your chrome profile
    options.add_argument(
        "user-data-dir=C:\\Users\\thiago.alves.EXTRABOM\\AppData\\Local\\Google\\Chrome\\User Data\Default")
    driver = webdriver.Chrome(chrome_options=options)
    driver.get("http://fluig.extrabom.com.br/portal/p/01/wcmuserpage")

    try:
        time.sleep(3)
        driver.find_element_by_xpath(
            '//*[@id="username"]').send_keys("processo.fluig")
        driver.find_element_by_xpath(
            '//*[@id="password"]').send_keys("")
        driver.find_element_by_xpath('//*[@id="submitLogin"]').click()
    except Exception:
        pass

    time.sleep(3)
    quantidade_pass = int(driver.find_element_by_xpath(
        '/html/body/div[3]/div[2]/div/div[1]/div/div[3]/div/div/div/div/div/div[2]/div/div/div/div[3]/p/strong').text[0:3]) // 100
    driver.find_element_by_xpath(
        '//*[@id="wcmid5_center"]/table/tbody/tr/td[8]/select').send_keys('100')

    csv = open('usuariofluig.csv', 'wt')
    csv.write("Login;Nome;Email;Status\n")

    for i in range(quantidade_pass+1):
        get_user(driver, csv)
        driver.find_element_by_xpath('//*[@id="next_wcmid5"]').click()
        time.sleep(5)

    csv.close()
    time.sleep(10)
    driver.close()


def get_user(driver, csv):
    time.sleep(1)
    elements = driver.find_elements_by_xpath(
        '/html/body/div[3]/div[2]/div/div[1]/div/div[3]/div/div/div/div/div/div[3]/div[2]/div[5]/div[3]/div/table/tbody')
    for element in elements:
        for i in range(6, len(element.find_elements_by_tag_name('td')), 1):
            text = element.find_elements_by_tag_name('td')[i].text
            print(text)
            if text != '':
                csv.write(text+";")
            else:
                csv.write("\n")
    csv.write("\n")


main()
