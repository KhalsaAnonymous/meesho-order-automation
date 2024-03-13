import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Initialize the Chrome WebDriver
#driver = webdriver.Chrome()
ch_options = Options()
ch_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=ch_options)
print("Opening Browser...")
# Read Excel file
input_file = 'input.xlsx'
output_file = 'output.xlsx'
df = pd.read_excel(input_file)
print("Reading data files...")
# # Set up web driver
# driver = webdriver.Chrome(ChromeDriverManager().install())

# Open website for login
print("Opening url...")
driver.get('https://supplier.meesho.com/panel/v3/new/root/login')  # Replace with the actual login page URL
time.sleep(1.5)
# Find the username and password fields and login button
username_field = driver.find_element(By.ID, "mui-1") # Replace with actual element ID
password_field = driver.find_element(By.ID, 'mui-2')  # Replace with actual element ID

login_button = driver.find_element(By.CLASS_NAME, 'MuiButton-containedPrimary')



# Provide login credentials and submit
username_field.send_keys('ID')
password_field.send_keys('Password')
login_button.click()
print("Loged in to the pannel...")

# Open website
print("Opening payments tab...")
time.sleep(1)
#driver.get('https://supplier.meesho.com/panel/v3/new/payouts/lsmfs/payments')  # Replace with the actual website URL for S99
driver.get('https://supplier.meesho.com/panel/v3/new/payouts/rthtj/payments')  # Replace with the actual website URL for JBV
time.sleep(0.6)
# Loop through Order IDs
for index, row in df.iterrows():
    time.sleep(1)
    order_id = row['Order_ID']
    # Find the search input field and submit button
    print("going to search.")
    search_input = driver.find_element(By.ID, 'filter-text-input')  # Replace with actual element ID


    # Clear the input field by selecting all and deleting
    search_input.send_keys(Keys.CONTROL + "a")  # Select all text
    search_input.send_keys(Keys.DELETE)  # Delete selected text

    # Enter order ID and submit
    # search_input.clear()
    search_input.send_keys(order_id)
    #search_button.click()
    search_input.send_keys(Keys.RETURN)  # This simulates pressing the Enter key

    try:
        time.sleep(2.5)

        # Extract required data
        print("Extracting required data..")
        time.sleep(0.50)
        order_details = driver.find_elements(By.TAG_NAME, 'td')
        sub_order_contribution = order_details[1].text
        net_order_amount = order_details[5].text
        # Update DataFrame
        print("writing data to file....")
        time.sleep(1.7)
        df.at[index, 'Sub Order Contribution'] = sub_order_contribution
        df.at[index, 'Net Order Amount'] = net_order_amount
    except IndexError:
        # Handle the case where order details are not found
        df.at[index, 'Sub Order Contribution'] = 'No Data'
        df.at[index, 'Net Order Amount'] = 'No Data'
    

# Write updated DataFrame to output Excel file
df.to_excel(output_file, index=False)
print("Data has been completed stored in output.xlsx")

# Close the browser
driver.quit()
print("close browser..")


