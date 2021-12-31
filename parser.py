from roster import *
from bs4 import BeautifulSoup
from datetime import datetime


def format_grade(grade):
    """Formats a grade into its abbreviated form

    Args:
        grade(str): The grade to format

    Returns:
        str: The grade in its abbreviated form

    """
    age_start = grade.find(' ') + 1
    age_end = grade.find(' ', age_start)
    section_start = grade.find(' ', age_end + 1) + 1
    section_end = grade.find(' ', section_start)
    return grade[age_start:age_end] + grade[section_start:section_end]


def get_grade(grade_html):
    """Gets the grade from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        str: The grade

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    grade = soup.find('h2', class_='sc-kEqYlL sc-1hg285i-0 eoUoDK hALyVo').text
    return format_grade(grade)


def format_date(date):
    """Formats a date into its desired format

    Args:
        date(str): The date to format

    Returns:
        str: The formatted date

    """
    return datetime.strptime(date, '%A, %d %B %Y').strftime('%d/%m/%Y')


def get_date(grade_html):
    """Gets the date from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        str: The date

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    date = soup.find('span', class_='sc-kEqYlL jndYxC').text
    return format_date(date)


def get_teams(grade_html):
    """Gets the teams from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(str): The teams

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    matches_element = soup.find('ul', class_='sc-10c3c88-4 iEXxNO')
    teams = matches_element.find_all('a', class_='sc-kEqYlL sc-10c3c88-13 gYjcIn johWCg')
    return [team.text for team in teams]


def get_time_location_htmls(grade_html):
    """Gets the HTML strings for the elements that hold the time and location information

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(str): The time/location HTML strings

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    time_location_elements = soup.find_all('div', class_='sc-10c3c88-15 ivbMVO')
    return [str(time_location_element) for time_location_element in time_location_elements]


def format_time(time):
    """Formats a time to remove the 'AM/PM' component

    Args:
        time(str): The time to format

    Returns:
        str: The formatted time

    """
    end = time.find(' ')
    return time[:end]


def get_times(grade_html):
    """Gets the times from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(str): The times

    """
    times = []

    time_location_htmls = get_time_location_htmls(grade_html)
    for time_location_html in time_location_htmls:
        soup = BeautifulSoup(time_location_html, 'html.parser')
        time = soup.find('span', class_='sc-kEqYlL kjKiYr').text
        times.append(format_time(time))

    return times


def get_location_court_strings(grade_html):
    """Gets the strings that contain the location and court information from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(str): The location/court strings

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    location_court_elements = soup.find_all('a', class_='sc-kEqYlL sc-10c3c88-20 bBbCEa kreAQ')
    return [location_court_element.text for location_court_element in location_court_elements]


def get_locations(location_court_strings):
    """Gets the locations from a list of location/court strings

    Args:
        location_court_strings(list(str)): The list of location/court strings

    Returns:
        list(str): The locations

    """
    locations = []

    for location_court_string in location_court_strings:
        end = location_court_string.find('/') - 1
        location = location_court_string[:end]
        locations.append(Location.from_official_name(location))

    return locations


def get_courts(location_court_strings):
    """Gets the courts from a list of location/court strings

    Args:
        location_court_strings(list(str)): The list of location/court strings

    Returns:
        list(str): The courts

    """
    courts = []

    for location_court_string in location_court_strings:
        search_term = 'Court'
        start = location_court_string.find(search_term) + len(search_term) + 1
        court = location_court_string[start:]
        courts.append(court)

    return courts


def create_matches(grade_html):
    """Gets the matches from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(str): The matches

    """
    matches = []

    teams = get_teams(grade_html)
    times = get_times(grade_html)

    location_court_strings = get_location_court_strings(grade_html)
    locations = get_locations(location_court_strings)
    courts = get_courts(location_court_strings)

    for i in range(len(teams) // 2):
        team1 = teams[i * 2]
        team2 = teams[i * 2 + 1]
        time = times[i]
        location = locations[i]
        court = courts[i]
        match = Match(team1, team2, time, location, court)
        matches.append(match)

    return matches


def create_round(grade_html):
    """Creates a Round object from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        Round: The created Round object

    """
    grade = get_grade(grade_html)
    date = get_date(grade_html)
    matches = create_matches(grade_html)
    round_ = Round(grade, date, matches)
    return round_


def create_roster(grade_htmls):
    """Creates a Roster object from a list of grade html strings

    Args:
        grade_htmls(list(str)): The grade HTML strings

    Returns:
        Roster: The created Roster object

    """
    rounds = []

    for grade_html in grade_htmls:
        round_ = create_round(grade_html)
        rounds.append(round_)

    roster = Roster(rounds)
    return roster
