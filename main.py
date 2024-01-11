import json
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from email_validator import validate_email, EmailNotValidError
import time
import pandas as pd

options = webdriver.EdgeOptions()
driver = webdriver.Edge(options=options)

driver.get("https://www.cambridgeinternational.org/why-choose-us/find-a-cambridge-school/")

selected_countries = ["Turkey", "Georgia", "Oman", "Egypt", "Qatar", "Turkmenistan", "Uzbekistan", "Kyrgyzstan", "Tajikistan", "Kazakhstan"]

def get_schoolsinfo():
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'result-container')))
    result_container = driver.find_element(By.CLASS_NAME, "result-container")  
    universities = result_container.find_elements(By.XPATH, "//tbody/tr/td/a")  

    schoolsinfo_list = []
    for university in universities:
        university_name = university.text
        university_link = university.get_attribute("href")
        schoolsinfo_list.append({"name": university_name, "link": university_link})

    return schoolsinfo_list

def extract_emails_from_page():
    email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
    page_source = driver.page_source
    emails = re.findall(email_pattern, page_source)
    return emails

def is_valid_email_format(text):
    try:
        # Validate the email address
        v = validate_email(text)

        # If no exception is raised, the email is valid
        return True

    except EmailNotValidError as e:
        # Handle the invalid email address case
        print(f"Error: {e}")
        return False

def fix_emails(data):
    for entry in data:
        if "emails" in entry:
            entry["emails"] = list(set(entry["emails"]))  
            entry["emails"] = [email for email in entry["emails"] if is_valid_email_format(email)]  
            if not entry["emails"]:
                del entry["emails"]  

def load_and_fix_json(filename):
    with open(filename, 'r', encoding='utf-8') as file:
        data = json.load(file)
        fix_emails(data)
    return data

def save_to_json(data, filename):
    with open(filename, 'a', encoding='utf-8') as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
        file.write(',\n')  

def save_to_excel(data, filename):
    try:
        existing_data = pd.read_excel(filename, engine='openpyxl')
    except FileNotFoundError:
        existing_data = pd.DataFrame()

    new_data = pd.DataFrame(data)
    final_data = pd.concat([existing_data, new_data], ignore_index=True)
    final_data.to_excel(filename, index=False, engine='openpyxl')

country_selector = Select(driver.find_element(By.ID, 'SelectedRegionId'))  
countries = country_selector.options

for country_index in range(1, len(countries)):
    country_selector.select_by_index(country_index)
    selected_country_option = country_selector.first_selected_option
    country = selected_country_option.get_attribute('value')

    if country == "Online" or country not in selected_countries:
        continue

    try:
        WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((By.ID, 'SelectedRegionId'), country))

        city_selector = Select(driver.find_element(By.ID, 'SelectedCity'))  
        cities = city_selector.options

        for city_index in range(1, len(cities)):
            city_selector.select_by_index(city_index)
            selected_city_option = city_selector.first_selected_option
            city = selected_city_option.get_attribute('value')

            try:
                WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((By.ID, 'SelectedCity'), city))

                search_button = driver.find_element(By.ID, "search")  
                search_button.click()
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)

                schoolsinfo = get_schoolsinfo()

                for info in schoolsinfo:
                    schoolsdata = {
                        "country": country,
                        "city": city,
                        "name": info['name'],
                        "link": info['link']
                    }

                    try:
                        driver.get(info['link'])
                        emails_on_page = extract_emails_from_page()
                        if emails_on_page:
                            schoolsdata["emails"] = [email for email in emails_on_page if is_valid_email_format(email)]
                        else:
                            schoolsdata["emails"] = []

                        save_to_json(schoolsdata, "schools-info-fix.json")
                        save_to_excel([schoolsdata], "schools-info.xlsx")
                        print(schoolsdata)

                    except Exception as e:
                        print(f"Error emails: {str(e)}")

                    finally:
                        driver.back()

            except Exception as e:
                print(f"Error city: {str(e)}")

    except Exception as e:
        print(f"Error country: {str(e)}")

        # Continue to the next country even if an error occurs
        continue

driver.quit()
