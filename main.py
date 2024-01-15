from selenium import webdriver
import time
import pickle as pkl
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import json

names = []
infos = {}

# Initialize the WebDriver
opt = webdriver.ChromeOptions()
opt.add_argument(f"user-data-dir=C:\\Users\\user\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
driverS = webdriver.ChromeService(executable_path=".\drivers\chromedriver.exe")
driver = webdriver.Chrome(service=driverS, options=opt)
# Navigate to the login page
driver.get('https://web.skype.com/?openPstnPage=true')

# Wait for manual login
input("Please log in manually in the opened browser window and then press Enter here...")

# Now that the user has logged in, you can navigate to the page you want to scrape
cookies = driver.get_cookies()
pkl.dump(cookies,open("cookies.pkl","wb"))



click_elems = []
####### Find Names

elements = driver.find_elements(By.XPATH, "//div[@aria-label and not(@title) and @role='button']")

# Print the outer HTML of each element
for element in elements:
    if not "group chat" in element.get_attribute('aria-label') and "chat" in element.get_attribute('aria-label'):
        nam = element.get_attribute('aria-label').split(",")[0]
        if nam in ["pinned","favorite"]:
            pass
        else:
            click_elems.append(element)
            names.append(nam)

print(names)
############ Find Ids

for elem in click_elems:
    elem.click()
    elements = driver.find_elements(By.XPATH, "//button[@aria-label][@role='button']")
    for element in elements:
        if element.get_attribute('aria-label').split(",")[0] in names and not element.get_attribute('aria-label').split(",")[0] in infos.keys():
            nm = element.get_attribute('aria-label').split(",")[0]
            infos[nm] = {}
            element.click()
            elmos = driver.find_elements(By.XPATH, "//button[@aria-label][@role='button']")
            for elm in elmos:
                elmm = elm.get_attribute('aria-label').split(", ")
                if elmm[0] == "Skype Name":
                    infos[nm]["SkypeID"]=elmm[1]
                elif elmm[0] == "Location":
                    infos[nm]["Location"]=",".join(elmm[1:]) 
                elif elmm[0] == "Mobile":
                    infos[nm]["Mobile"]=elmm[1]
            button = driver.find_element(By.XPATH, "//button[@role='button'][@title='Close user profile'][@aria-label='Close user profile']")
            button.click()
            break           
    time.sleep(1)
    json.dump(infos,open("info.json","w"),indent=4)

def json_to_csv(json_data, file_path):
    """
    Convert a JSON dictionary to a CSV file.

    :param json_data: Dictionary containing the JSON data.
    :param file_path: Path where the CSV file will be saved.
    """
    import csv

    # Determine all unique keys in the JSON
    all_keys = set()
    for user_info in json_data.values():
        for key in user_info.keys():
            all_keys.add(key)

    # Write to CSV
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        
        # Write header
        header = ["Username"] + list(all_keys)
        writer.writerow(header)

        # Write data rows
        for username, user_info in json_data.items():
            row = [username] + [user_info.get(key, "") for key in all_keys]
            writer.writerow(row)

    return file_path

json_to_csv(infos,"info.csv")

# # Close the WebDriver
driver.quit()
