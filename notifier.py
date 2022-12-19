import os
os.environ["PATH"] += ":"+os.getcwd()

import re
import time
import getpass
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from threading import Thread
from bs4 import BeautifulSoup
from pushbullet import Pushbullet
pb = Pushbullet(open("pb.key", "r").read())

regex_hour = '\d{2}:\d{2}' # match hh:mm
regex_day = '\d{2}\/\d{2}' # match mm/dd
delay = 5 # seconds

options = Options()
options.headless = True
email = input("Corporate account: ")
corp_domain = email.split('@')[-1]
password = getpass.getpass()

driver_outlook = webdriver.Firefox(options=options)
driver_teams = webdriver.Firefox(options=options)

def login_microsoft(driver):
    WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div > input[type=email]')))

    selector = driver.find_element(By.CSS_SELECTOR, 'div > input[type=email]')
    selector.send_keys(email)
    selector.send_keys(Keys.RETURN)

    WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div > a#idA_PWD_ForgotPassword')))

    selector = driver.find_element(By.CSS_SELECTOR, 'div > input[type=password]')
    selector.send_keys(password)
    selector.send_keys(Keys.RETURN)

def teams_scraper(unread_teams):
    chat_scroll = 'ul[data-tid="chat-list-tree"]'
    try:
        myElem = WebDriverWait(driver_teams, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, chat_scroll)))
        # print('hit__')
    except Exception as e:
        # print(e)
        driver_teams.delete_all_cookies()
        driver_teams.refresh()
        driver_teams.get("https://teams.microsoft.com/")
        
        login_microsoft(driver_teams)

        selector = WebDriverWait(driver_teams, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'li > button[aria-label^=Chat]')))
        selector.click()

        myElem = WebDriverWait(driver_teams, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, chat_scroll)))

    soup = BeautifulSoup(myElem.get_attribute('innerHTML'), "html.parser")
    for div in soup.select('div > div:has(> a[class*="left-rail-unread"])'):
        title,__,body = div.select('span[class*="single-line-truncation"]')
        title = "Teams - "+title.getText()
        body = body.getText()
        unread_teams.add((title,body))
    return unread_teams

def outlook_scraper(unread_outlook):
    email_scroll = 'div[role="listbox"] > div.customScrollBar'
    try:
        myElem = WebDriverWait(driver_outlook, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, email_scroll)))
        # print('hit_')
    except Exception as e:
        # print(e)
        driver_outlook.delete_all_cookies()
        driver_outlook.refresh()
        driver_outlook.get(f"http://outlook.com/owa/{corp_domain}/")

        login_microsoft(driver_outlook)

        myElem = WebDriverWait(driver_outlook, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, email_scroll)))

    soup = BeautifulSoup(myElem.get_attribute('innerHTML'), "html.parser")

    for div in soup.select('div > div[aria-label^=Unread]'):
        preview = div['aria-label'][7:]
        try:
            title, body = re.split(regex_hour, preview)
        except:
            try:
                title, body = re.split(regex_day, preview)
            except:
                title = preview
                body = ''

        title = ''.join(c for c in title if c.isprintable()).strip()
        body = ''.join(c for c in body if c.isprintable()).strip()
        if title =='': title = '-'
        if body =='': body = '-'

        title = "Outlook - "+title
        unread_outlook.add((title,body))
    return unread_outlook

while True:
    unread_teams = set()
    unread_outlook = set()
    teams_thread = Thread(target=teams_scraper, args=(unread_teams,))
    outlook_thread = Thread(target=outlook_scraper, args=(unread_outlook,))

    teams_thread.start()
    outlook_thread.start()

    teams_thread.join()
    outlook_thread.join()
    unread = unread_teams.union(unread_outlook)

    notifications = pb.get_pushes()
    pushed = set([(b['title'], b['body']) for b in notifications])
    to_unpush = pushed - unread
    to_push = unread - pushed

    for n in notifications:
        if (n['title'], n['body']) in to_unpush:
            pb.delete_push(n["iden"])

    for t,b in to_push:
        pb.push_note(t, b)
        
    time.sleep(2)