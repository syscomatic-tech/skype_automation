from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import traceback
import pandas as pd
import json
import time

infos = []

# Initialize the WebDriver
# opt = webdriver.ChromeOptions()
# opt.add_argument(f"user-data-dir=.\Default")
driverS = webdriver.ChromeService(executable_path=".\drivers\chromedriver.exe")
driver = webdriver.Chrome(service=driverS)
fpath = input("please enter the filepath: ")
# Navigate to the login page
driver.get('https://web.skype.com/?openPstnPage=true')

def expand(driver):
    while True:
        try:
            # Locate the div with the specified data-text-as-pseudo-element attribute
            div_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@role='group' and @tabindex='-1' and @aria-label='Skype directory']//div[@data-text-as-pseudo-element='More']"))
            )

            # Click on the located div
            div_element.click()

        except Exception as e:
            break

def parse_info(input_string):
    """
    Parses the input string into a dictionary with name, mutual connections, Skype name, and location.

    Args:
    input_string (str): A string in the format "Name, X mutual connection(s), Skype Name: skype_name, location: location"

    Returns:
    dict: A dictionary with keys 'name', 'mutual_connections', 'skype_name', and 'location'.
    """
    try:
        # Split the string by commas to separate the different parts
        parts = input_string.split(',')

        # Extracting individual pieces of information
        name = parts[0].strip()
        if "mutual connection" in input_string:
            mutual_connections = input_string.strip().split(" mutual connection")[0].split(",")[1]
            if mutual_connections != "Web Research":
                mutual_connections = int(mutual_connections)
            else:
                mutual_connections = 0

        else:
            mutual_connections = 0
        skype_name = input_string.strip().split('Skype Name: ')[1].split(",")[0]
        if "location" in input_string:
            location = input_string.strip().split('location: ')[1]
        else:
            location = ""

        # Constructing the dictionary
        info_dict = {
            'name': name,
            'mutual_connections': mutual_connections,
            'skype_name': skype_name,
            'location': location
        }

        return info_dict
    except Exception as e:
        trace = traceback.format_exc()

        print("An error occured: ",e)
        print("Traceback:", trace)

        info_dict = {
            'name': "",
            'mutual_connections': "",
            'skype_name': "",
            'location': ""
        }

        return info_dict 


def get_contacts(driver):
    xpath = "//div[@role='group' and @tabindex='-1' and @aria-label='Skype directory']//div[@role='listitem'][@aria-label]"
    elements = driver.find_elements(By.XPATH,xpath)
    aria_labels = [parse_info(element.get_attribute('aria-label')) for element in elements]
    infos.extend(aria_labels)
    with open("infos.json","w",encoding="utf-8") as fl:
        json.dump(infos,fl,indent=4)

def json_to_unique_csv(json_data, file_path):
    """
    Converts a JSON object to a CSV file with unique rows.

    Args:
    json_data (list): A list of dictionaries representing JSON objects.
    file_path (str): The file path where the CSV file will be saved.
    """
    # Convert JSON to DataFrame
    df = pd.DataFrame(json_data)

    # Remove duplicate rows
    unique_df = df.drop_duplicates()

    # Convert DataFrame to CSV and save to the specified file path
    unique_df.to_csv(file_path, index=False)
    print(f"CSV file saved to {file_path}")
# Wait for manual login
input("Please log in manually in the opened browser window and then press Enter here...")

try:
    # Click on the specified div
    div_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@aria-label='People, groups, messages, web' and @role='button']"))
    )
    div_element.click()

    # Locate the input element
    input_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Search Skype' and @size='1' and @aria-label='Search Skype' and @type='text']"))
    )

    # Type each letter from 'a' to 'z', clearing after each
    for char in range(ord('a'), ord('z') + 1):
        input_element.send_keys(chr(char))
        expand(driver)
        time.sleep(1)  # Wait for 1 second after typing each letter
        get_contacts(driver)
        input_element.send_keys(Keys.BACKSPACE)
    json_to_unique_csv(infos,fpath)
except Exception as e:
    print("Error occurred: ", e)
finally:
    driver.quit()
