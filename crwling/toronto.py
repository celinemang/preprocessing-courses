import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Setup WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)
driver.get('https://ttb.utoronto.ca/')

# Click the initial dropdown
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "division"))
).click()

# Wait for options and click them
options_text = [
    'Faculty of Applied Science & Engineering',
    'Faculty of Arts and Science',
    'Faculty of Information',
    'Faculty of Kinesiology and Physical Education',
    'Faculty of Music',
    'John H. Daniels Faculty of Architecture, Landscape, & Design',
    'University of Toronto Mississauga',
    'University of Toronto Scarborough'
]

for text in options_text:
    option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//span[contains(text(), '{text}')]"))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", option)
    driver.execute_script("arguments[0].click();", option)
    WebDriverWait(driver, 2).until(
        lambda driver: driver.execute_script('return document.readyState') == 'complete'
    )

# Click the specified element before clicking additional options
specified_element_xpath = "/html/body/app-root/main/app-dashboard/app-search/div/div[1]/div/app-select-criteria[2]/div/div[2]"
specified_element = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, specified_element_xpath))
)
driver.execute_script("arguments[0].scrollIntoView(true);", specified_element)
driver.execute_script("arguments[0].click();", specified_element)

# Click the additional options
additional_options_xpath = [
    "/html/body/app-root/main/app-dashboard/app-search/div/div[1]/div/app-select-criteria[2]/div/div[3]/app-ttb-option[1]/span",
    "/html/body/app-root/main/app-dashboard/app-search/div/div[1]/div/app-select-criteria[2]/div/div[3]/app-ttb-option[2]/span",
    "/html/body/app-root/main/app-dashboard/app-search/div/div[1]/div/app-select-criteria[2]/div/div[3]/app-ttb-option[3]/span",
    "/html/body/app-root/main/app-dashboard/app-search/div/div[1]/div/app-select-criteria[2]/div/div[3]/app-ttb-option[4]/span"
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

# Click the search button
search_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary"))
)
driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
driver.execute_script("arguments[0].click();", search_button)
time.sleep(10)
courses = []
page_number = 1


        
for i in range(1, 3):  # 각 페이지에서 실제 존재하는 코스 수만큼 반복
    # try:
        # Course_Name and Section
        course_xpath = f"/html/body/app-root/main/app-dashboard/div/app-result/div[2]/app-course[{i}]"
        course_element = driver.find_element(By.XPATH, course_xpath)
        
        # 열기 버튼 클릭
        open_button_xpath = f"/html/body/app-root/main/app-dashboard/div/app-result/div[2]/app-course[{i}]/div/div/h3/button"
        open_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, open_button_xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", open_button)
        driver.execute_script("arguments[0].click();", open_button)
        time.sleep(1)  # Ensure the content loads
        
        # More Course Information 버튼 클릭
        try:
            more_info_button_xpath = f"/html/body/app-root/main/app-dashboard/div/app-result/div[2]/app-course[{i}]/div/div/div/div/div[5]/div/div[11]/span"
            more_info_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, more_info_button_xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", more_info_button)
            driver.execute_script("arguments[0].click();", more_info_button)
        except:
            pass  # If the more info button is not found, continue without clicking
        
        course_name_section = course_element.find_element(By.XPATH, "./div/div/h3/button/span").text.strip()

        # Extract course name and section
        if ':' in course_name_section:
            section_part, course_name_part = course_name_section.split(':', 1)
            section = section_part.strip()
            course_name = course_name_part.strip()
        else:
            section = ""
            course_name = course_name_section

        # 다른 정보 수집
        def get_text_or_empty(xpath):
            try:
                element_text = driver.find_element(By.XPATH, xpath).text
                return "" if element_text in ["", "TBA"] else element_text
            except:
                return ""

        credit_xpath = f"/html/body/app-root/main/app-dashboard/div/app-result/div[2]/app-course[{i}]/div/div/div/div/div[5]/div/div[11]/span"
        day_time_xpath = f"/html/body/app-root/main/app-dashboard/div/app-result/div[2]/app-course[{i}]/div/div/div/div/div[6]/div/app-course-section/div/div[1]/div[1]/div/div/span"
        room_xpath = f"/html/body/app-root/main/app-dashboard/div/app-result/div[2]/app-course[{i}]/div/div/div/div/div[6]/div/app-course-section/div/div[1]/div[2]/div/div/div"
        professor_xpath = f"/html/body/app-root/main/app-dashboard/div/app-result/div[2]/app-course[{i}]/div/div/div/div/div[6]/div/app-course-section/div/div[1]/div[3]/div/span"

        credit = get_text_or_empty(credit_xpath)
        day_time = get_text_or_empty(day_time_xpath)
        room = get_text_or_empty(room_xpath)
        professor = get_text_or_empty(professor_xpath)

        course_info = {
            "Course Name": course_name,
            "Section": section,
            "Credit": credit,
            "Day & Time": day_time,
            "Room": room,
            "Professor": professor
        }
        courses.append(course_info)
        print(course_info)  # Print course information

   

# # 다음 페이지로 이동
# try:
#     pagination_xpath = "/html/body/app-root/main/app-dashboard/div/app-result/div[2]/ngb-pagination[2]/ul"
#     pagination_elements = driver.find_elements(By.XPATH, pagination_xpath + "/li")
    
#     if not pagination_elements:
#         # Check for the alternate pagination
#         pagination_xpath_alt = "/html/body/app-root/main/app-dashboard/div/app-result/div[2]/ngb-pagination[1]/ul/li"
#         pagination_elements = driver.find_elements(By.XPATH, pagination_xpath_alt)
    
#     if len(pagination_elements) <= 1:
#         print("No more pages available.")
        
    
#     next_page_element = pagination_elements[-1]  # 마지막 li 요소를 선택
#     next_page_button = next_page_element.find_element(By.TAG_NAME, 'a')

#     if 'disabled' in next_page_element.get_attribute('class'):
#         print("Last page reached.")
        
#     else:
#         next_page_button.click()

# except Exception as e:
#     print(f"No more pages available or error navigating to next page: {e}")
    

page_number += 1
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, f"/html/body/app-root/main/app-dashboard/div/app-result/div[2]/app-course[1]/div/div/h3/button/span"))
)



# Save the results to an Excel file
df = pd.DataFrame(courses)
df.to_excel("courses.xlsx", index=False)
print("Data saved to courses.xlsx")

# WebDriver 닫기
driver.quit()
