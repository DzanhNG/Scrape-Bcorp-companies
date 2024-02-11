from ast import While
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import datetime
import tkinter as tk
from tkinter import FIRST, messagebox
import requests


class WebScraper:
    def __init__(self):
        # Initialize the webdriver
        self.firefox_options = Options()
        self.firefox_options.headless = True
        self.driver = webdriver.Firefox(options=self.firefox_options)
        self.actions = ActionChains(self.driver)
        self.data_export = []

    def close_browser(self):
        self.driver.quit()

    def navigate_to_url(self, url):
        self.driver.get(url)
        sleep(10)  # Adjust sleep as needed

    def get_data_from_Firstpage(self):
        # Now that the content is loaded, get the HTML
        print('Render the dynamic content to static HTML')
        html = self.driver.page_source       
        print(' Parse the static HTML')
        soup = BeautifulSoup(html, "html.parser")
        for li_tag in soup.find_all('li', class_='ais-Hits-item'):
            linkx = li_tag.find('a')['href']
            link = 'https://www.bcorporation.net' + linkx

            ## For each Profile:
            response = requests.get(link)
            if response.status_code == 200:
                soup_2 = BeautifulSoup(response.text, 'html.parser')
                # Check if the <i> tag with class "fa fa-users" exists
                name = soup_2.find('div', class_ ="col-start-5 col-end-12 py-4").find('h1').text
                # Extract "Headquarters" information
                headquarters_span = soup_2.find('span', text='Headquarters')
                headquarters_info = headquarters_span.find_next('p').text.strip() if headquarters_span else None

                # Extract "Industry" information
                industry_span = soup_2.find('span', text='Industry')
                industry_info = industry_span.find_next('p').text.strip() if industry_span else None

                print("name:", name)
                print("Headquarters:", headquarters_info)
                print("Industry:", industry_info)




                self.data_export.append({'Company Name':name, 'URL': link, 'Headquarters':headquarters_info, 'Industry':industry_info  })



            else:
                print(f"Failed to retrieve the page. Status code: {response.status_code}")


    def export_to_excel(self, file_name='output.xlsx'):
        # Create a DataFrame from the list of dictionaries
        df = pd.DataFrame(self.data_export)

        # Export DataFrame to Excel
        df.to_excel(file_name, index=False)
        print(f'Data exported to {file_name}')

# Example usage:
if __name__ == "__main__":
    ##Filter the option in Marketplace and copy the url
    scraper = WebScraper()
    url1 = "https://www.bcorporation.net/en-us/find-a-b-corp/"
    scraper.navigate_to_url(url1)
    scraper.get_data_from_Firstpage()

    for i in range(2,10):
        url2 = "https://www.bcorporation.net/en-us/find-a-b-corp/?page=" + str(i)
        scraper.navigate_to_url(url2)
        scraper.get_data_from_Firstpage()
    
    # Export the data to Excel after the loop
    scraper.export_to_excel()

    scraper.close_browser()