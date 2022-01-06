from bs4 import BeautifulSoup
from datetime import datetime
from roster import *


def _get_date(grade_html):
    """Gets the date from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        datetime: The date

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    date_string = soup.find('span', class_='sc-kEqYlL jndYxC').text
    return datetime.strptime(date_string, '%A, %d %B %Y')


def _format_grade(grade):
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


def _get_grade(grade_html):
    """Gets the grade from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        str: The grade

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    grade = soup.find('h2', class_='sc-kEqYlL sc-1hg285i-0 eoUoDK hALyVo').text
    return _format_grade(grade)


def _get_teams(grade_html):
    """Gets the teams from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(str): The teams

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    matches_element = soup.find('ul', class_='sc-10c3c88-4 iEXxNO')
    all_teams = matches_element.find_all('a', class_='sc-kEqYlL sc-10c3c88-13 gYjcIn johWCg')

    # Filter out forfeited matches
    teams = []
    for i in range(len(all_teams) // 2):
        skip = False
        teams_to_add = []
        for j in range(2):
            team = all_teams[i * 2 + j]
            if team.find_next_sibling('span', class_='sc-kEqYlL kTltqj') is not None:
                skip = True
                break
            teams_to_add.append(team)

        if skip:
            continue

        teams.extend(teams_to_add)

    return [team.text for team in teams]


def _get_time_location_htmls(grade_html):
    """Gets the HTML strings for the elements that hold the time and location information

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(str): The time/location HTML strings

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    time_location_elements = soup.find_all('div', class_='sc-10c3c88-15 ivbMVO')
    return [str(time_location_element) for time_location_element in time_location_elements]


def _get_times(grade_html):
    """Gets the times from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(datetime): The times

    """
    times = []

    time_location_htmls = _get_time_location_htmls(grade_html)
    for time_location_html in time_location_htmls:
        soup = BeautifulSoup(time_location_html, 'html.parser')
        time_string = soup.find('span', class_='sc-kEqYlL kjKiYr').text
        time = datetime.strptime(time_string, '%I:%M %p')
        times.append(time)

    return times


def _get_location_court_strings(grade_html):
    """Gets the strings that contain the location and court information from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(str): The location/court strings

    """
    soup = BeautifulSoup(grade_html, 'html.parser')
    location_court_elements = soup.find_all('a', class_='sc-kEqYlL sc-10c3c88-20 bBbCEa kreAQ')
    return [location_court_element.text for location_court_element in location_court_elements]


def _get_locations(location_court_strings):
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


def _get_courts(location_court_strings):
    """Gets the courts from a list of location/court strings

    Args:
        location_court_strings(list(str)): The list of location/court strings

    Returns:
        list(Court): The courts

    """
    courts = []

    for location_court_string in location_court_strings:
        search_term = 'Court'
        start = location_court_string.find(search_term) + len(search_term) + 1
        court = int(location_court_string[start:])
        courts.append(Court.from_num(court))

    return courts


def _create_matches(grade_html):
    """Gets the matches from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        list(str): The matches

    """
    matches = []

    grade = _get_grade(grade_html)
    teams = _get_teams(grade_html)
    times = _get_times(grade_html)

    location_court_strings = _get_location_court_strings(grade_html)
    locations = _get_locations(location_court_strings)
    courts = _get_courts(location_court_strings)

    for i in range(len(teams) // 2):
        team1 = teams[i * 2]
        team2 = teams[i * 2 + 1]
        time = times[i]
        location = locations[i]
        court = courts[i]
        match = Match(grade, team1, team2, time, location, court)
        matches.append(match)

    return matches


def _create_round(grade_html):
    """Creates a Round object from a grade html string

    Args:
        grade_html(str): The grade HTML string

    Returns:
        Round: The created Round object

    """
    matches = _create_matches(grade_html)
    round_ = Round(matches)
    return round_


def create_roster(grade_htmls):
    """Creates a Roster object from a list of grade html strings

    Args:
        grade_htmls(list(str)): The grade HTML strings

    Returns:
        Roster: The created Roster object

    """
    rounds = []

    # Check if the date is correct
    date = _get_date(grade_htmls[0])
    for grade_html in grade_htmls:
        round_ = _create_round(grade_html)
        rounds.append(round_)

    roster = Roster(date, rounds)
    return roster
