import re

from requests_html import HTMLSession
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

_DOMAIN = 'https://www.playhq.com'
_COMPETITIONS_URL = _DOMAIN + '/basketball-victoria/org/southern-basketball-association/e1cbc3e3'

# The timeout value in seconds for loading javascript on a webpage
_JS_LOAD_TIMEOUT = 10


def _get_html(url):
    """Gets the HTML of a webpage

    Args:
        url(str): The URL of the webpage

    Returns:
        str: The HTML of the webpage

    """
    session = HTMLSession()
    response = session.get(url)
    return response.content


def _get_grades_url(competitions_html):
    """Gets the URL of the Saturday Junior Domestic grades page

    Args:
        competitions_html(str): The HTML of the competitions page

    Returns:
        str: The URL of the Saturday Junior Domestic grades page

    """
    soup = BeautifulSoup(competitions_html, 'html.parser')
    ref = soup.find_all('a', href=re.compile('junior-domestic'))[-1]['href']
    return _DOMAIN + ref


def _get_grade_urls(grades_html):
    """Gets the list of urls that correspond to the fixtures for all Saturday grades

    Args:
        grades_html(str): The HTML of the grades page

    Returns:
        list(str): The list of urls that correspond to the fixtures for all Saturday grades

    """
    soup = BeautifulSoup(grades_html, 'html.parser')
    all_grades_element = soup.find('ul', class_='sc-12ty7r5-0 hrILMC sc-1vy00ws-2 dBMkSW')
    saturday_grade_elements = all_grades_element.find_all('a', href=re.compile('saturday'))
    return [_DOMAIN + grade['href'] for grade in saturday_grade_elements]


def _wait_for_html(driver):
    """Waits for a webpage to load the required javascript, and then returns its HTML

    Args:
        driver(WebDriver): The web driver to use

    Returns:
        str: The HTML of the webpage

    """
    loaded_condition = EC.presence_of_element_located((By.CSS_SELECTOR, '.sc-10c3c88-5.gdnNoD'))
    WebDriverWait(driver, _JS_LOAD_TIMEOUT).until(loaded_condition)
    return driver.page_source


def _get_htmls_with_js(urls):
    """Gets the HTML after the javascript has loaded from a list of provided URLs
       If any given webpage or the javascript in the webpage cannot load, it will retry once

    Args:
        urls(list(str)): The URLs of the webpages

    Returns:
        list(str): The list of HTML strings

    """
    htmls = []

    driver = webdriver.Chrome()

    # Positions the window off the screen, so it's not visible
    driver.set_window_position(-2000, 0)

    # Loads each url in a tab
    driver.get(urls[0])
    for i, url in zip(range(2, len(urls) + 1), urls[1:]):
        driver.execute_script(f'window.open(\'about:blank\', \'{i}\');')
        driver.switch_to.window(str(i))
        driver.get(url)

    # Iterates through the tabs, loads the javascript and then saves the HTML
    handles = driver.window_handles
    for i, handle in zip(range(len(handles)), handles):
        driver.switch_to.window(handle)

        try:
            htmls.append(_wait_for_html(driver))
        except TimeoutException:
            driver.get(urls[i])
            try:
                htmls.append(_wait_for_html(driver))
            except TimeoutException:
                driver.quit()
                return None

    driver.quit()
    return htmls


def get_all_grade_htmls():
    """Gets the HTML of all required grade pages

    Returns:
        list(str): The list of grade page HTML strings

    """
    competitions_html = _get_html(_COMPETITIONS_URL)
    grades_url = _get_grades_url(competitions_html)
    grades_html = _get_html(grades_url)
    grade_urls = _get_grade_urls(grades_html)
    grade_htmls = _get_htmls_with_js(grade_urls)
    return grade_htmls
