from enum import Enum


class Location(Enum):
    """Represents the location that a match can take place at

    """
    KING_CLUB = 1,
    PARKDALE = 2,
    MENTONE_GRAMMAR = 3,
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


class Match:
    """Represents a match between two teams

    """
    def __init__(self, team1, team2, time, location, court):
        self.team1 = team1
        self.team2 = team2
        self.time = time
        self.location = location
        self.court = court


class Round:
    """Represents a round of matches

    """
    def __init__(self, grade, date, matches):
        self.grade = grade
        self.date = date
        self.matches = matches


class Roster:
    """Represents a roster of rounds

    """
    def __init__(self, rounds):
        self.rounds = rounds

    def __str__(self):
        """Converts a roster into a string, mainly for debugging purposes

        Returns:
            str: the Roster as a string

        """
        string = '\n\n'
        for round_ in self.rounds:
            string += f'Grade - {round_.grade}:\n'
            string += f'Date - {round_.date}:\n'
            string += f'Matches:\n'
            for match in round_.matches:
                string += f'\t{match.team1} vs {match.team2}:\n'
                string += f'\t\tTime - {match.time}\n'
                string += f'\t\tLocation - {match.location}\n'
                string += f'\t\tCourt - {match.court}\n'
            string += '\n'
        string += '\n\n'

        return string
