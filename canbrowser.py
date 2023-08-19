import os
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import time
import undetected_chromedriver as uc

class CanBrowser:
    def __init__(self):
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.parent_dir = os.path.dirname(self.current_dir)
        # chrome profile PATH 
        self.profile_path = os.path.join(self.parent_dir, 'chrome-profile')
        # ChromeDriver PATH
        self.chrome_driver_path = os.path.join(self.parent_dir, 'chrome-driver', 'chromedriver.exe')
        # Chrome-win PATH
        self.chrome_path = os.path.join(self.parent_dir, 'chrome-win', 'chrome.exe')
        self.driver = None

    def create_webdriver(self, profile_name):
        chrome_options = Options()
        chrome_options.binary_location = self.chrome_path
        profile_dir = os.path.join(self.profile_path, profile_name)
        chrome_options.add_argument("--user-data-dir=" + profile_dir)
        service = Service(self.chrome_driver_path)
        self.driver = uc.Chrome(service=service, options=chrome_options)
    
    def create_profile(self, profile_id):
        profile_name = f"profile_{profile_id}"
        profile_dir = os.path.join(self.profile_path, profile_name)
        os.makedirs(profile_dir, exist_ok=True)
        self.create_webdriver(profile_name)
        
    def open_profile(self, profile_id):
        self.create_profile(profile_id)
    
    def close_profile(self):
        if self.driver:
            self.driver.quit()

    def delete_profile(self, profile_id):
        profile_name = f"profile_{profile_id}"
        profile_dir = os.path.join(self.profile_path, profile_name)
        if os.path.exists(profile_dir):
            shutil.rmtree(profile_dir)

if __name__ == "__main__":
    profile1 = CanBrowser()
    profile1.open_profile(2)
    time.sleep(2)
    if profile1.driver:
        profile1.driver.get("https://www.google.com")
    
    while True:
        time.sleep(1)
        


