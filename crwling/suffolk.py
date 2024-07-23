import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime

# Setup WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)
driver.get('https://www.suffolk.edu/coursesearch/courses')

# Click Academic Period -> Fall 2024
additional_options_xpath = [
    "/html/body/div[1]/div/section/aside/div/div[1]/div/ul/li[6]/div",
    "/html/body/div[1]/div/section/aside/div/div[1]/div/ul/li[6]/ul/li/ul/li[1]/label/span[1]/span"
]

for xpath in additional_options_xpath:
    option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", option)
    driver.execute_script("arguments[0].click();", option)
    WebDriverWait(driver, 2).until(
        lambda driver: driver.execute_script('return document.readyState') == 'complete'
    )


def get_text_or_empty(xpath):
    try:
        element_text = driver.find_element(By.XPATH, xpath).text
        return "" if element_text in ["", "TBA"] else element_text
    except:
        return ""

def convert_to_24h_format(day_time_str):
    try:
        parts = day_time_str.split('|')
        days = parts[0].strip().split('_')
        time_range = parts[1].strip()
        start_time, end_time = time_range.split('-')
        start_time_24 = datetime.strptime(start_time.strip(), '%I:%M %p').strftime('%H:%M')
        end_time_24 = datetime.strptime(end_time.strip(), '%I:%M %p').strftime('%H:%M')
        converted_parts = [f"{day}/{start_time_24}-{end_time_24}" for day in days]
        return ','.join(converted_parts)
    except Exception as e:
        print(f"Error converting time: {e}")
        return day_time_str

page_number = 1
total_pages = 153

courses = []

for page_number in range(1, total_pages + 1):
    for i in range(1, 10):  # Adjust the range according to the actual number of courses per page
        try:
            # Course Name and Section
            course_path = f"/html/body/div[1]/div/section/section/main/div[1]/div/div/div[{i}]/div/div[1]/div[1]/div[1]/b"
            section_path1 = f"/html/body/div[1]/div/section/section/main/div[1]/div/div/div[{i}]/div/div[1]/div[1]/div[2]/div/button[1]/span"
            section_path2 = f"/html/body/div[1]/div/section/section/main/div[1]/div/div/div[{i}]/div/div[1]/div[1]/div[2]/div/button[2]/span"

            # Collect additional information
            credit_xpath = f"/html/body/div[1]/div/section/section/main/div[1]/div/div/div[{i}]/div/div[1]/div[2]/span[3]/span"
            day_time_xpath = f"/html/body/div[1]/div/section/section/main/div[1]/div/div/div[{i}]/div/div[2]/div/div/div[2]/div[1]/div[2]"
            professor_xpath = f"/html/body/div[1]/div/section/section/main/div[1]/div/div/div[{i}]/div/div[2]/div/div/div[2]/div[1]/div[4]/div/span"
            
            credit = get_text_or_empty(credit_xpath).split('|')[0].strip()
            day_time = get_text_or_empty(day_time_xpath)
            professor = get_text_or_empty(professor_xpath)
            course_name = get_text_or_empty(course_path)
            section1 = get_text_or_empty(section_path1)
            section2 = get_text_or_empty(section_path2)

            section = section1 + section2

            # Convert day & time to 24-hour format
            day_time_24h = convert_to_24h_format(day_time)

            # Append course information
            course_info = {
                "Section": section,
                "Course Name": course_name,
                "Credit": credit,
                "Professor": professor,
                "Time/Date": day_time_24h,
                "Room": "",  # Placeholder for Room
                "Course Code": ""  # Placeholder for Course Code
            }
            courses.append(course_info)
            print(course_info)  # Print course information
        except Exception as e:
            print(f"Error processing course {i} on page {page_number}: {e}")
            continue

    # Move to the next page
    try:
        next_page_xpath = f"/html/body/div[1]/div/section/section/main/div[2]/div/ul/li[{page_number + 2}]/a"
        next_page_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, next_page_xpath))
        )
        next_page_element.click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"/html/body/div[1]/div/section/section/main/div[2]/div/ul/li[{page_number + 2}]"))
        )
    except Exception as e:
        print(f"No more pages available or error navigating to next page: {e}")
        break

# Save the results to an Excel file
df = pd.DataFrame(courses)
df.to_excel("suffolk_Fall2024.xlsx", index=False)
print("Data saved to suffolk_Fall2024.xlsx")

# Close WebDriver
driver.quit()
