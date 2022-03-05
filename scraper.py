import re
import os

import chromedriver_autoinstaller

from datetime import datetime
from requests_html import HTMLSession
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from exception import RoundNotFoundException
from window import update_progress, update_driver
from export import GRADES_TO_SKIP

# Import a Windows specific constant if the current platform is Windows
if os.name == 'nt':
    from subprocess import CREATE_NO_WINDOW

# URL constants
_DOMAIN = 'https://www.playhq.com'
_COMPETITIONS_URL = _DOMAIN + '/basketball-victoria/org/southern-basketball-association/e1cbc3e3'

# The timeout value in seconds for loading a page (excluding js)
_PAGE_LOAD_TIMEOUT = 15

# The timeout value in seconds for loading javascript on a webpage
_JS_LOAD_TIMEOUT = 5

# Define how much progress is made by each section
_SHALLOW_SCRAPE_LENGTH = 30
_OPENING_TABS_LENGTH = 60
_DEEP_SCRAPE_LENGTH = 5


def _get_html(url):
    """Gets the HTML of a webpage

    Args:
        url(str): The URL of the webpage

    Returns:
        str: The HTML of the webpage

    """
    session = HTMLSession()
    response = session.get(url)
    return str(response.content)


def _url_to_grade(grade_url):
    """Extracts the grade from a grade URL

    Args:
        grade_url(str): The grade URL to extract the grade from

    Returns:
        str: The extracted grade

    """
    search_term = 'saturday'
    start = grade_url.find(search_term) + len(search_term) + 1
    end = grade_url.find('/', start)
    grade_section = grade_url[start:end]
    grade = grade_section.replace('-', ' ').title()
    return grade


def _get_grade_htmls(grade_urls):
    """Gets the HTML of a list of grade URLs

    Args:
        grade_urls(list(str)): The URLs of the grade webpages

    Returns:
        list(str): The HTML of the webpages

    """
    htmls = []
    for i, grade_url in zip(range(len(grade_urls)), grade_urls):
        update_info = f'Shallow scraping the \'{_url_to_grade(grade_url)}\' page...'
        update_progress(update_info, int((_SHALLOW_SCRAPE_LENGTH / len(grade_urls)) * (i + 1)))
        htmls.append(_get_html(grade_url))

    return htmls


def _get_grades_url(competitions_html):
    """Gets the URL of the Saturday Junior Domestic grades page

    Args:
        competitions_html(str): The HTML of the competitions page

    Returns:
        str: The URL of the Saturday Junior Domestic grades page

    """
    soup = BeautifulSoup(competitions_html, 'html.parser')
    ref = soup.find_all('a', href=re.compile('junior-domestic'))[0]['href']
    return _DOMAIN + ref


def _get_grade_urls(grades_html):
    """Gets the list of urls that correspond to the fixtures for all Saturday grades

    Args:
        grades_html(str): The HTML of the grades page

    Returns:
        list(str): The list of URLs that correspond to the fixtures for all Saturday grades

    """
    soup = BeautifulSoup(grades_html, 'html.parser')
    all_grades_element = soup.find('ul', class_='sc-12ty7r5-0 hrILMC sc-1vy00ws-2 dBMkSW')
    saturday_grade_elements = all_grades_element.find_all('a', href=re.compile('saturday'))
    return [_DOMAIN + grade['href'] for grade in saturday_grade_elements]


def _get_grade_url_by_round(grade_html, round_index):
    """Gets the url that corresponds to the fixtures for a specific round

    Args:
        grade_html(str): The HTML of the grade page
        round_index(int): The round index

    Returns:
        str: The URL that corresponds to the fixtures for a specific round

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    round_elements = soup.find_all('a', class_='sc-2zsuyh-3 kQmXVu')
    round_element = round_elements[round_index]
    grade_url = _DOMAIN + round_element['href']
    return grade_url


def _wait_for_html(driver):
    """Waits for a webpage to load the required javascript, and then returns its HTML

    Args:
        driver(WebDriver): The driver to use

    Returns:
        str: The HTML of the webpage

    """
    loaded_condition = EC.presence_of_element_located((By.CSS_SELECTOR, '.sc-10c3c88-5.gdnNoD'))
    WebDriverWait(driver, _JS_LOAD_TIMEOUT).until(loaded_condition)
    return driver.page_source


def _wait_for_html_error(driver):
    """Waits for a webpage to load an error, and then returns its HTML

    Args:
        driver(WebDriver): The driver to use

    Returns:
        str: The HTML of the webpage

    """
    loaded_condition = EC.presence_of_element_located((By.CSS_SELECTOR, '.n806zu-0.eOOEPz'))
    WebDriverWait(driver, _JS_LOAD_TIMEOUT).until(loaded_condition)
    return driver.page_source


def _wait_for_html_unconfirmed(driver):
    """Waits for a webpage to load an unconfirmed message, and then returns its HTML

    Args:
        driver(WebDriver): The driver to use

    Returns:
        str: The HTML of the webpage

    """
    loaded_condition = EC.presence_of_element_located((By.CSS_SELECTOR, '.n806zu-0.kxpuUz.sc-10c3c88-18.dsJxqP'))
    WebDriverWait(driver, _JS_LOAD_TIMEOUT).until(loaded_condition)
    return driver.page_source


def _is_saturday_match(grade_html):
    """Checks if a Saturday match has actually been scheduled for Saturday

    Args:
        grade_html(str): The HTML of the grade page

    Returns:
        bool: True if the match is scheduled for Saturday, otherwise False

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    date_element = soup.find('span', class_='sc-kEqYlL jndYxC')
    unconfirmed_element = soup.find('div', class_='n806zu-0 kxpuUz sc-10c3c88-18 dsJxqP')
    if date_element and not unconfirmed_element:
        date_text = date_element.text
        return 'saturday' in date_text.lower()

    grade_text = soup.find('h2', class_='sc-kEqYlL sc-1hg285i-0 eoUoDK hALyVo').text
    age_start = grade_text.find(' ') + 1
    age_end = grade_text.find(' ', age_start)
    section_start = grade_text.find(' ', age_end + 1) + 1
    grade_text = grade_text[age_start:age_end] + grade_text[section_start:]
    GRADES_TO_SKIP.append(grade_text)
    return False


# noinspection all
def _get_htmls_with_js(urls):
    """Gets the HTML after the javascript has loaded from a list of provided URLs
       If any given webpage or the javascript in the webpage cannot load, it will retry once

    Args:
        urls(list(str)): The URLs of the webpages

    Returns:
        list(str): The list of HTML strings

    """
    htmls = []

    # Initialise chrome driver
    update_progress('Downloading Chrome driver...', _SHALLOW_SCRAPE_LENGTH)
    driver_path = chromedriver_autoinstaller.install()

    # Hide the Chrome driver console if on Windows
    if os.name == 'nt':
        service = Service(driver_path)
        service.creationflags = CREATE_NO_WINDOW
        driver = webdriver.Chrome(service=service)
    else:
        driver = webdriver.Chrome()

    driver.set_page_load_timeout(_PAGE_LOAD_TIMEOUT)
    update_driver(driver)

    # Position the window off the screen, so it's not visible
    driver.set_window_position(-10000, 0)

    # A list of grades to display while loading
    grades = [_url_to_grade(urls[0])]

    # Loads each url in a tab
    update_info = f'Opening the \'{grades[-1]}\' page...'
    update_progress(update_info, _SHALLOW_SCRAPE_LENGTH + (_OPENING_TABS_LENGTH // len(urls)))
    driver.get(urls[0])
    for i, url in zip(range(2, len(urls) + 1), urls[1:]):
        grades.append(_url_to_grade(url))
        update_info = f'Opening the \'{grades[-1]}\' page...'
        update_progress(update_info, _SHALLOW_SCRAPE_LENGTH + int((_OPENING_TABS_LENGTH / len(urls)) * i))

        driver.execute_script(f'window.open(\'about:blank\', \'{i}\');')
        driver.switch_to.window(str(i))
        driver.get(url)

    # Iterate through the tabs, load the javascript and then save the HTML
    update_progress('Deep scraping the opened pages...', _SHALLOW_SCRAPE_LENGTH + _OPENING_TABS_LENGTH)
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
                driver.get(urls[i])
                try:
                    htmls.append(_wait_for_html_error(driver))
                except TimeoutException:
                    driver.get(urls[i])
                    try:
                        htmls.append(_wait_for_html_unconfirmed(driver))
                    except TimeoutException as e:
                        driver.quit()
                        update_driver(None)
                        raise e

    driver.quit()
    update_driver(None)
    return list(filter(lambda html: _is_saturday_match(html), htmls))


def _get_current_date(grade_html):
    """Gets the date from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        str: The date

    """
    search_term = 'provisionalDate'
    start = grade_html.find('"current":true')
    start = grade_html.find(search_term, start) + len(search_term) + 1
    start = grade_html.find('"', start) + 1
    end = grade_html.find('"', start)
    date_string = grade_html[start:end]
    date_string = datetime.strptime(date_string, '%Y-%m-%d').strftime('%d/%m/%Y')
    return date_string


def _get_grade_dates(grade_html):
    """Gets the dates for each round from a grade HTMl string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(str): The list of dates

    """
    dates = []
    search_term = 'provisionalDate'
    date_matches = [match.start() for match in re.finditer(search_term, grade_html)]
    for date_match in date_matches:
        start = date_match + len(search_term) + 1
        start = grade_html.find('"', start) + 1
        end = grade_html.find('"', start)
        date_string_to_add = grade_html[start:end]
        date_string_to_add = datetime.strptime(date_string_to_add, '%Y-%m-%d').strftime('%d/%m/%Y')
        dates.append(date_string_to_add)

    return dates


def _transform_grade_urls(grade_urls, date_string):
    """Make sure the grade urls have the correct date

    Args:
        grade_urls: The grade URLs to sanitise

    Returns:
        list(str): The sanitised grade URLs

    """
    # Get the HTML for each URL
    grade_htmls = _get_grade_htmls([f'{grade_url}/R1' for grade_url in grade_urls])
    dates = [_get_grade_dates(grade_html) for grade_html in grade_htmls]

    # Verify that there is at least one grade playing on the date specified
    round_not_found = True
    for date_strings in dates:
        if date_string in date_strings:
            round_not_found = False
            break

    if round_not_found:
        raise RoundNotFoundException(date_string)

    round_indices = []
    index = 0
    for dates_ in dates:
        if date_string not in dates_:
            grade_htmls.pop(index)
            continue

        date_index = dates_.index(date_string)
        round_indices.append(date_index)
        index += 1

    return [_get_grade_url_by_round(grade_htmls[i], round_indices[i]) for i in range(len(grade_htmls))]


def get_all_grade_htmls(date_string):
    """Gets the HTML of all required grade pages

    Returns:
        list(str): The list of grade page HTML strings

    """
    competitions_html = _get_html(_COMPETITIONS_URL)
    grades_url = _get_grades_url(competitions_html)
    grades_html = _get_html(grades_url)
    grade_urls = _get_grade_urls(grades_html)
    grade_urls = _transform_grade_urls(grade_urls, date_string)
    grade_htmls_with_js = _get_htmls_with_js(grade_urls)
    return grade_htmls_with_js
