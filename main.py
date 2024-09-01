import os
import time
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import Workbook

# Set up the Chrome WebDriver
os.environ['PATH'] += r"C:\Python"
driver = webdriver.Chrome()

# Create a directory to save images
if not os.path.exists('ebay_images'):
 os.makedirs('ebay_images')
 
# Set up the Excel workbook and sheet
wb = Workbook()
ws = wb.active
ws.title = "ebay Items"
ws.append(["Name", "Price", "Condition", "Item Location", "Image Filename", "Product URL"])

# Define the search term and URL
search_term = "exercise equipment"
url = f"https://www.ebay.com/sch/i.html?_from=R40&_nkw={search_term}" 

# Open eBay and search for the term
driver.get(url)
time.sleep(3)  # wait for the page to load

# Set a counter to limit the number of items scraped
max_items = 500
items_collected = 0
page = 1

while items_collected < max_items:
  driver.get(f"{url}&_pgn={page}")
  time.sleep(3)  # wait for the page to load
  
  # Find all product containers on the page
  product_containers = driver.find_elements(By.CSS_SELECTOR,'ul.srp-results.srp-list.clearfix li')
  
  for idx, container in enumerate(product_containers):
    if items_collected >= max_items:
      break
    
    try:
     # Extract product name
     name = container.find_element(By.CSS_SELECTOR,'div.s-item__title').text.strip()
     
     # Extract product price
     price = container.find_element(By.CSS_SELECTOR,'span.s-item__price').text.strip()
     
     # Extract product condition
     condition = container.find_element(By.CSS_SELECTOR,'div.s-item__subtitle').text.strip()
     
     # Extract item location
     item_location = container.find_element(By.CSS_SELECTOR,'span.s-item__location.s-item__itemLocation').text.strip()
     
     # Extract product URL
     product_url = container.find_element(By.CSS_SELECTOR,'a.s-item__link').get_attribute("href")
     
     # Extract and download the image
     image_url = container.find_element(By.CSS_SELECTOR,'div.s-item__image-wrapper.image-treatment img').get_attribute('src')
     image_filename = f"ebay_images/{search_term}_{items_collected + 1}.jpg"
     
     # Download and save the image
     img_data = requests.get(image_url).content
     with open(image_filename,'wb') as handler:
       handler.write(img_data)
       
     # Save product details in Excel
     ws.append([name,price,condition,item_location,image_filename,product_url])
     items_collected += 1
     print(f"Collected item {items_collected}")

    except Exception as e:
     print(f"Failed to process item {idx + 1}: {e}")
     
  # Move to the next page
  page += 1
  time.sleep(2)
# Save the workbook
wb.save("ebay_items.xlsx")

# Close the browser
driver.quit()

print("Scraping completed and data saved to ebay_items.xlsx.")  