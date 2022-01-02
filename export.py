from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Font, PatternFill, Alignment, Border, Side
from datetime import datetime
from roster import Location, Court

_ORDINAL_INDICATORS = ('th', 'st', 'nd', 'rd')
_DATE_FORMAT = '%d/%m/%Y'
_TIME_FORMAT = '%I:%M'
_ROW_HEIGHT = 17
_DATE_CELL = 'C4'
_FIRST_TABLE_ROW = 6
_COURT_COLUMN = _TIME_COLUMN = 'B'
_LOCATION_COLUMN = _TEAM_1_COLUMN = 'C'
_TEAM_2_COLUMN = 'D'
_GRADE_COLUMN = 'E'
_REFEREE_COLUMN_1 = 'F'
_REFEREE_COLUMN_2 = 'G'


def _ordinal(date):
    """Gets the correct ordinal indicator for a date

    Args:
        date(datetime): The date to get the ordinal indicator for

    Returns:
        str: The correct ordinal indicator

    """
    day = int(date.strftime('%d'))

    if day % 10 in (1, 2, 3) and day not in (11, 12, 13):
        return _ORDINAL_INDICATORS[day % 10]

    return _ORDINAL_INDICATORS[0]


def create_excel(roster, template_location, save_location):
    """Creates an Excel document from a roster

    Args:
        roster(Roster): The roster to export
        template_location(str): The location of the template Excel document
        save_location(str): The location of the folder to save the Excel document to

    """
    wb = load_workbook(template_location)
    data = roster.to_dictionary()

    # Define global style components
    bold_font = Font(bold=True)
    center_alignment = Alignment(horizontal='center')

    # Define table style components
    table_side = Side(style='thin', color="000000")
    table_border = Border(left=table_side, top=table_side, right=table_side, bottom=table_side)

    # Define time style components
    time_font = Font(bold=True, color='FFFFFF')
    time_fill = PatternFill(bgColor='000000', fill_type='solid')

    # Define and add named styles
    wb.add_named_style(NamedStyle(name='date', font=bold_font, alignment=center_alignment))
    wb.add_named_style(NamedStyle(name='court', font=bold_font, alignment=center_alignment))
    wb.add_named_style(NamedStyle(name='location', font=bold_font))
    wb.add_named_style(NamedStyle(name='time', font=time_font, alignment=center_alignment, fill=time_fill))
    wb.add_named_style(NamedStyle(name='team', font=bold_font, border=table_border))
    wb.add_named_style(NamedStyle(name='grade', font=bold_font, border=table_border, alignment=center_alignment))
    wb.add_named_style(NamedStyle(name='referee', border=table_border))

    date = data['Date']
    for location in Location:
        ws = wb[str(location)]
        row = _FIRST_TABLE_ROW

        # Fill in the date cell
        date_cell = ws[_DATE_CELL]
        date_cell.value = date.strftime(_DATE_FORMAT).lstrip('0')
        date_cell.style = 'date'

        # Fill in court tables
        for court in Court:
            matches = data[location][court]
            if len(matches) == 0:
                continue

            # Fill in court cell
            court_cell = ws[f'{_COURT_COLUMN}{row}']
            court_cell.value = f'Crt {court.value}'
            court_cell.style = 'court'

            # Fill in location cell
            location_cell = ws[f'{_LOCATION_COLUMN}{row}']
            location_cell.value = str(location)
            location_cell.style = 'location'

            # Increase height of court row
            ws.row_dimensions[row].height = _ROW_HEIGHT

            row += 1

            # Fill in match rows
            for match in matches:
                # Fill in match time cell
                match_time_cell = ws[f'{_TIME_COLUMN}{row}']
                match_time_cell.value = match.time.strftime(_TIME_FORMAT).lstrip('0')
                match_time_cell.style = 'time'

                # Fill in match team 1 cell
                match_team_1_cell = ws[f'{_TEAM_1_COLUMN}{row}']
                match_team_1_cell.value = match.team1
                match_team_1_cell.style = 'team'

                # Fill in match team 2 cell
                match_team_2_cell = ws[f'{_TEAM_2_COLUMN}{row}']
                match_team_2_cell.value = match.team2
                match_team_2_cell.style = 'team'

                # Fill in match grade cell
                match_grade_cell = ws[f'{_GRADE_COLUMN}{row}']
                match_grade_cell.value = match.grade
                match_grade_cell.style = 'grade'

                # Format referee cells
                referee_cells = ws[f'{_REFEREE_COLUMN_1}{row}':f'{_REFEREE_COLUMN_2}{row}'][0]
                for referee_cell in referee_cells:
                    referee_cell.style = 'referee'

                # Increase height of match row
                ws.row_dimensions[row].height = _ROW_HEIGHT

                row += 1

    # Set the first worksheet as active
    wb.active = 0

    # Save and close the worksheet
    filename = date.strftime('%d{} %b %Y').format(_ordinal(date)).lstrip('0') + '.xlsx'
    wb.save(f'{save_location}/{filename}')
    wb.close()


def update_excel(roster, excel_location):
    """Updates an Excel document with a new roster

    Args:
        roster(Roster): The new roster to update the Excel document with
        excel_location(str): The location of the Excel document to update

    """
    pass
