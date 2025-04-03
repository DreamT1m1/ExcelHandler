from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.ie.webdriver import WebDriver
from selenium.webdriver.support.wait import WebDriverWait

driver = webdriver.Chrome()


driver.get("http://tehnika/hooldus/toolaud")

button = WebDriverWait(driver, 2).until(
    EC.element_to_be_clickable((By.ID, "sisselogimisenupp"))
)
button.click()
print("Вход выполнен\n")

input_field = driver.find_element(By.CSS_SELECTOR, "input.form-control")

value = "516566"

input_field.clear()

input_field.send_keys(value)

input_field.send_keys(Keys.RETURN)

time.sleep(2)

is_complex = False

try:
    table = driver.find_element(By.ID, "sugseade-tab")
    rows = table.find_elements(By.TAG_NAME, "tr")
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        if value in [i.text for i in cells]:
            row.click()
            print("Click done!")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    try:
        complex = driver.find_element(By.ID, "komplekt")
    except NoSuchElementException:
        pass
    else:
        is_complex = True

    info = driver.find_elements(By.TAG_NAME, "b")

    for el in info:
        parent = el.find_element(By.XPATH, "..")
        full_text = parent.text.strip()
        data = dict(filter(lambda x: len(x) == 2, [tuple(i.split(": ")) for i in full_text.split("\n")]))

        if not is_complex:
            if data["RP-kood"] == value:
                if "Seadme nr" in data:
                    print("Seadme nr", data["Seadme nr"])
                else:
                    print("No device code found")
                    print("Seerianr", data["Seerianr"])
        else:
            print("Complex of devices")
            print(f"Complex s/n: {data['Seadme nr']}")

        break

except NoSuchElementException:
    print("No device found")

driver.quit()
