from enum import Enum
from collections import defaultdict


class Location(Enum):
    """Represents the location that a match can take place at

    """
    KING_CLUB = 1
    PARKDALE = 2
    MENTONE_GRAMMAR = 3
    MENTONE_GIRLS = 4

    @classmethod
    def from_official_name(cls, string):
        """Converts a location's official name into a Location enum value

        Args:
            string(str): The location's official name

        Returns:
            Location: The resulting Location enum value

        """
        if string == 'Sandringham Family Leisure Centre':
            return cls.KING_CLUB
        if string == 'Parkdale Secondary College':
            return cls.PARKDALE
        if string == 'Mentone Grammar School':
            return cls.MENTONE_GRAMMAR
        if string == 'Mentone Girls Secondary College':
            return cls.MENTONE_GIRLS

    def __str__(self):
        """Converts a Location enum value into a readable string

        Returns:
            str: the Location enum value as a readable string

        """
        return self.name.replace('_', ' ').title()


class Court(Enum):
    """Represents the court that a match can take place on

    """
    COURT_1 = 1
    COURT_2 = 2
    COURT_3 = 3
    COURT_4 = 4

    @classmethod
    def from_num(cls, num):
        """Converts a court's number into a Court enum value

        Args:
            num(int): The court's number

        Returns:
            Court: The resulting Court enum value

        """
        if num == cls.COURT_1.value:
            return cls.COURT_1
        if num == cls.COURT_2.value:
            return cls.COURT_2
        if num == cls.COURT_3.value:
            return cls.COURT_3
        if num == cls.COURT_4.value:
            return cls.COURT_4

    def __str__(self):
        return self.name.replace('_', ' ').title()


class Match:
    """Represents a match between two teams

    """
    def __init__(self, grade, team1, team2, time, location, court):
        self.grade = grade
        self.team1 = team1
        self.team2 = team2
        self.time = time
        self.location = location
        self.court = court


class Round:
    """Represents a round of matches

    """
    def __init__(self, matches):
        self.matches = matches


class Roster:
    """Represents a roster of rounds

    """
    def __init__(self, date, rounds):
        self.date = date
        self.rounds = rounds

    @classmethod
    def _insert_match_in_order(cls, match, matches):
        """Inserts a Match object into a list of matches in order based on the match time

        Args:
            match(Match): The Match object to insert
            matches(list(Match)): The matches

        """
        if len(matches) == 0:
            matches.append(match)
        else:
            for i, curr_match in zip(range(len(matches)), matches):
                if match.time < curr_match.time:
                    matches.insert(i, match)
                    break
                else:
                    if i == len(matches) - 1:
                        matches.append(match)

    def to_dictionary(self):
        """Converts the roster object into a dictionary

        Returns:
            defaultdict: The roster object as a dictionary

        """
        data = defaultdict(lambda: defaultdict(list))

        data['Date'] = self.date
        for round_ in self.rounds:
            for match in round_.matches:
                Roster._insert_match_in_order(match, data[match.location][match.court])

        return data
