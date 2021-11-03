import os
import logging
from typing import List
from time import sleep
from configparser import ConfigParser
from functools import wraps

import pandas as pd
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from openpyxl import load_workbook
from selenium.webdriver.remote.webelement import WebElement


class Robot:
    def __init__(self, output_file: str):
        self.filename = output_file
        self.browser_lib = Selenium()
        self.pdf_files = []

    def open_the_website(self, url: str):
        """Opens the website and click Dive In button"""

        dl_path = f'{os.path.abspath(os.getcwd())}/output/'
        prefs = {"download.default_directory": dl_path, "safebrowsing.enabled": "false"} 
        
        self.browser_lib.open_available_browser(url, use_profile=True, preferences=prefs)
        self.browser_lib.click_link('#home-dive-in')

    def process_single_agency(self, agencies: list, required_agency: str) -> WebElement:
        """Loads the whole investment table of certain agency and return it"""
        
        # Finding agency throughout all the agencies and go to its page 
        for agency in agencies:
            name = agency.find_element_by_class_name('h4.w200').text
            if name == required_agency:
                break
        agency.click()

        # Wait until table is loaded and choose 'show all' option
        logging.info("Waiting until Investment Table is loaded..")
        option_css = "css:select.form-control:nth-child(1) > option:nth-child(4)"
        self.browser_lib.wait_until_element_is_visible(option_css, timeout=60)
        self.browser_lib.find_element(option_css).click()

        # Wait until the whole table will be loaded
        self.browser_lib.wait_until_element_is_not_visible(
            "css:#investments-table-object_paginate > span:nth-child(3) > a:nth-child(2)", timeout=60)

        table = self.browser_lib.find_element('css:#investments-table-object')
        logging.info("Investment Table was parsed!")
        return table

    def get_pdfs(self, table: WebElement):
        """Extract links from table and download pdf files from them"""

        pdf_links = [el.get_attribute('href') for el in table.find_elements_by_tag_name('a')]
        for link in pdf_links:
            self.browser_lib.go_to(link)

            btn_css = "css:#business-case-pdf > a:nth-child(1)"
            # Wait until loads the download button and click it
            self.browser_lib.wait_until_element_is_visible(btn_css, timeout=30)
            self.browser_lib.find_element(btn_css).click()

            # Wait until generating pdf
            logging.info('Generating PDF..')
            self.browser_lib.wait_until_element_is_not_visible("css:#business-case-pdf > span:nth-child(2)", timeout=30)

            # Wait until file is downloaded
            self.wait_download(f'{os.path.abspath(os.getcwd())}/output/', 30)
            logging.info("PDF successfully downloaded!")
            self.pdf_files.append(link.split('/')[-1]+'.pdf')

    def wait_download(self, path: str, timeout: int):
        """Waits until files are downloaded"""

        logging.info('Downloading PDF..')
        wait = True
        seconds = 0
        while wait and seconds < timeout:
            sleep(1)
            files = os.listdir(path)
            wait = False
            for fname in files:
                if fname.endswith('.crdownload'):
                    wait = True
            
            seconds += 1

    def write_investments_to_excel(self, table):
        """Writes investments table to the second excel sheet by converting html to pandas dataframe"""

        pd_df = pd.read_html(table.get_attribute("outerHTML"))
        path = f"output/{self.filename}"

        wb = load_workbook(path)
        writer = pd.ExcelWriter(path, engine='openpyxl')
        writer.book = wb
        pd_df[0].to_excel(writer, sheet_name="Individual Investments", index=False)
        writer.save()

    def get_agencies(self) -> list:
        """Loads all the agencies and returns list of them"""

        # Waits until agencies will be loaded
        self.browser_lib.wait_until_element_is_visible(
            "css:#agency-tiles-widget > div:nth-child(1) > div:nth-child(1) > div:nth-child(1)", timeout=60)
        agencies = self.browser_lib.find_element(
            "css:#home-dive-in > div:nth-child(1) > div:nth-child(2)").find_elements_by_class_name("col-sm-12")

        # Waits until all the text will be loaded
        self.browser_lib.wait_until_element_is_visible(agencies[0].find_element_by_class_name('h4.w200'), timeout=60)
        return agencies

    def process_data_to_write(self, agencies: list) -> List[List[str]]:
        """Get name and spent amount of each agency preparing it for writing"""

        agencies_to_write = []
        for agency in agencies:
            name = agency.find_element_by_class_name('h4.w200').text
            amount = agency.find_element_by_class_name('h1.w900').text
            agencies_to_write.append([name, amount])
        return agencies_to_write

    def write_agencies_to_excel(self, agencies: list):
        """Writes name and spent amount to excel"""

        lib = Files()
        if not os.path.exists('output'):
            os.mkdir('output')

        lib.create_workbook(f'./output/{self.filename}')
        lib.rename_worksheet("Sheet", "Agencies")
        # Add headers
        lib.set_cell_value(1, 1, "Name")
        lib.set_cell_value(1, 2, "Spending")
        # Write main data
        lib.append_rows_to_worksheet(self.process_data_to_write(agencies), "Agencies")
        logging.info("Agencies written to excel")
        lib.save_workbook()

    def get_data_from_pdf(self, pdf_name: str):
        """Extract `Name of this Investment` and `Unique Investment Identifier (UII)` from PDF"""

        logging.disable(logging.INFO)  # Disable logging
        pdf = PDF()
        text = pdf.get_text_from_pdf(f"./output/{pdf_name}")

        # Extract values from raw text
        data_list = text[1].split('1. Name of this Investment:')[-1].split('2. Unique Investment Identifier (UII):')
        name = data_list[0].strip()
        uii = data_list[1].split('Section')[0].strip()

        logging.disable(logging.DEBUG)  # Enable logging back
        return {'Name': name, 'UII': uii}

    def get_data_from_table(self):
        """Generator which returns Inv Name and UII from each row"""

        lib = Files()
        lib.open_workbook(f'./output/{self.filename}')
        data = lib.read_worksheet("Individual Investments")

        # Gets accurate letter of required columns 
        name_letter = [l for l in data[0] if data[0][l] == 'Investment Title'][0]
        uii_letter = [l for l in data[0] if data[0][l] == 'UII'][0]
        for row in data[1:]:
            yield {'Name': row[name_letter], 'UII': row[uii_letter]}
        

    def compare_data(self):
        """Comparing data from each pdf and table of certain agency"""

        for pdf_name in self.pdf_files:
            pdf_data = self.get_data_from_pdf(pdf_name)
            name_from_pdf = pdf_data['Name']
            uii_from_pdf = pdf_data['UII']
            logging.info(f"Starting comparing data on file {pdf_name}.")
            for i, data_dict in enumerate(self.get_data_from_table()):
                name, uii = data_dict['Name'], data_dict['UII']
                if name == name_from_pdf and uii == uii_from_pdf:
                    logging.info(f"MATCH FOUND ON ROW {i+1}")
            

def main():
    # Read data from `settings.ini`
    try:
        config = ConfigParser()
        config.read('settings.ini')
        agency = config.get('Settings', 'agency')
        filename = config.get('Settings', 'filename') 
    except Exception:
        raise Exception
    
    robot = Robot(output_file=filename)

    try:
        robot.open_the_website("https://itdashboard.gov/")

        agencies = robot.get_agencies()  # Get list of all the agencies
        robot.write_agencies_to_excel(agencies)  # Writes name, spending to excel

        table = robot.process_single_agency(agencies, agency)
        robot.write_investments_to_excel(table)  # Writes Investment table to the same excel
        
        robot.browser_lib.set_download_directory(f"{os.path.abspath(os.getcwd())}/output/", download_pdf=True) 
        robot.get_pdfs(table)  # Download PDFs of asked agency
        robot.compare_data()  # Compare data from pdf and excel
    finally:
        robot.browser_lib.close_all_browsers()


if __name__ == "__main__":
    logging.basicConfig(format=u'%(filename)s [LINE:%(lineno)d] #%(levelname)-8s [%(asctime)s]  %(message)s',
                    level=logging.INFO)
    log = logging.getLogger(__name__)
    main()
