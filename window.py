import PySimpleGUI as sg
import os
import sys
import calendar

from datetime import datetime, timedelta

# Window constants
WINDOW_WIDTH = 760
MAIN_WINDOW_CREATE_HEIGHT = 250
MAIN_WINDOW_UPDATE_HEIGHT = 180
WINDOW_TAB_GROUP_DIFF = 115
TAB_GROUP_CREATE_HEIGHT = MAIN_WINDOW_CREATE_HEIGHT - WINDOW_TAB_GROUP_DIFF
TAB_GROUP_UPDATE_HEIGHT = MAIN_WINDOW_UPDATE_HEIGHT - WINDOW_TAB_GROUP_DIFF
PROGRESS_WINDOW_HEIGHT = 100
PROGRESS_WINDOW_HEIGHT_WITH_OPTIONS = 150
MAIN_COLUMN_KEY = '-MAIN COLUMN-'
TEMPLATE_DOCUMENT_KEY = '-TEMPLATE DOCUMENT TEXT-'
OUTPUT_FOLDER_KEY = '-OUTPUT FOLDER TEXT-'
CALENDAR_KEY = '-CALENDAR TEXT-'
UPDATE_DOCUMENT_KEY = '-UPDATE DOCUMENT TEXT-'
TAB_GROUP_KEY = '-TAB GROUP-'
PROCESS_BUTTON_KEY = '-PROCESS BUTTON-'
PROGRESS_COLUMN_KEY = '-PROGRESS COLUMN-'
PROGRESS_TEXT_KEY = '-PROGRESS TEXT-'
PROGRESS_BAR_KEY = '-PROGRESS BAR-'
PROGRESS_OPTIONS_KEY = '-PROGRESS OPTIONS COLUMN-'
PROGRESS_UTILITY_BUTTON_KEY = '-PROGRESS UTILITY BUTTON-'
PROGRESS_RESTART_BUTTON_KEY = '-PROGRESS RESTART BUTTON-'
PROGRESS_EXIT_BUTTON_KEY = '-PROGRESS EXIT BUTTON-'
THREAD_PROGRESS_EVENT = '-THREAD PROGRESS-'
THREAD_DRIVER_EVENT = '-THREAD DRIVER-'
ERROR_EVENT = '-ERROR-'
WINDOW_EXIT_EVENT = 'Exit'
CREATE_TAB = 'Create'
UPDATE_TAB = 'Update'
EXPLORER_UTILITY = 'Show in Explorer'
RETRY_UTILITY = 'Retry'
PROGRESS_BAR_COLOUR = ('green', 'white')
PROGRESS_BAR_ERROR_COLOUR = ('#8b0000', '#8b0000')


def _next_saturday(as_string):
    """Get the date of next Saturday

    Returns:
        tuple: The date of next saturday

    """
    now = datetime.now()
    saturday_delta = timedelta((calendar.SATURDAY - now.weekday()) % 7)
    saturday_datetime = now + saturday_delta
    if as_string:
        return saturday_datetime.strftime('%d/%m/%Y')

    return saturday_datetime.month, saturday_datetime.day, saturday_datetime.year


def change_window_height(height):
    """Changes the window height, making sure it doesn't move

    Args:
        height(int): The new window height

    """
    window_x, window_y = WINDOW.current_location()
    WINDOW.size = (WINDOW_WIDTH, height)
    WINDOW.move(window_x, window_y)


def switch_to_progress_layout():
    """Switch to the progress layout

    """
    # Change which layout and elements are visible
    main_column = WINDOW[MAIN_COLUMN_KEY]
    progress_column = WINDOW[PROGRESS_COLUMN_KEY]
    main_column.update(visible=False)
    progress_column.update(visible=True)

    # Change the window height
    change_window_height(PROGRESS_WINDOW_HEIGHT)


def toggle_progress_options(error=False):
    """Toggles the visibility of the progress options

    Args:
        error(bool): Whether an error occurred that halted the progress bar

    """
    progress_options_column = WINDOW[PROGRESS_OPTIONS_KEY]
    showing = progress_options_column.visible
    if showing:
        progress_options_column.update(visible=False)
        change_window_height(PROGRESS_WINDOW_HEIGHT)
        return

    utility_button = WINDOW[PROGRESS_UTILITY_BUTTON_KEY]
    utility_button.update(text=RETRY_UTILITY if error else EXPLORER_UTILITY)
    progress_options_column.update(visible=True)
    change_window_height(PROGRESS_WINDOW_HEIGHT_WITH_OPTIONS)


def update_progress(info, value):
    """Sends an event to the window with progress information

    Args:
        value(int): The new progress value
        info(str): Information on what is currently processing

    """
    WINDOW.write_event_value(THREAD_PROGRESS_EVENT, (info, value))


def update_driver(driver):
    """Sends an event to the window with a driver

    Args:
        driver(WebDriver): The driver to send

    """
    WINDOW.write_event_value(THREAD_DRIVER_EVENT, driver)


def update_error(error_msg):
    """Sends an event to the window with an error message

    Args:
        error_msg(str): The error message to send

    """
    WINDOW.write_event_value(ERROR_EVENT, error_msg)


# Window layouts and elements
_TEMPLATE_ROW_1 = [
    sg.Text('Template Document:', size=(18, 1)),
    sg.In(
        os.getcwd().replace('\\', '/') + '/resources/template.xlsx',
        key=TEMPLATE_DOCUMENT_KEY,
        size=(60, 1),
        disabled=True,
        enable_events=True
    ),
    sg.FileBrowse(file_types=(('Excel Documents', '*.xlsx'),))
]
_OUTPUT_ROW_1 = [
    sg.Text('Output Folder:', size=(18, 1)),
    sg.In(key=OUTPUT_FOLDER_KEY, size=(60, 1), disabled=True, enable_events=True),
    sg.FolderBrowse()
]
_CALENDAR_ROW_1 = [
    sg.Text('', expand_x=True),
    sg.In(
        _next_saturday(True),
        key=CALENDAR_KEY,
        size=(11, 1),
        justification='center',
        disabled=True
    ),
    sg.CalendarButton(
        'Choose Date',
        target=CALENDAR_KEY,
        format='%d/%m/%Y',
        default_date_m_d_y=_next_saturday(False),
        no_titlebar=False
    ),
    sg.Text('', expand_x=True),
]
_UPDATE_ROW_2 = [
    sg.Text('Document to Update:', size=(18, 1)),
    sg.In(key=UPDATE_DOCUMENT_KEY, size=(60, 1), disabled=True, enable_events=True),
    sg.FileBrowse(file_types=(('Excel Documents', '*.xlsx'),))
]
_MAIN_COLUMN_1 = [
    _TEMPLATE_ROW_1,
    _OUTPUT_ROW_1,
    _CALENDAR_ROW_1
]
_MAIN_COLUMN_2 = [
    _UPDATE_ROW_2
]
_TAB_GROUP_ROW = [
    sg.TabGroup([[
        sg.Tab('Create', [[sg.Column(_MAIN_COLUMN_1, pad=15)]], expand_x=True, expand_y=True),
        sg.Tab('Update', [[sg.Column(_MAIN_COLUMN_2, pad=15)]], expand_x=True)
    ]], key=TAB_GROUP_KEY, tab_location='center', enable_events=True)
]
_PROCESS_ROW = [
    sg.Button('Process', key=PROCESS_BUTTON_KEY, size=(10, 2), disabled=True)
]
_MAIN_COLUMN = [
    _TAB_GROUP_ROW,
    _PROCESS_ROW
]
_MAIN_LAYOUT = [
    sg.Column(
        _MAIN_COLUMN,
        key=MAIN_COLUMN_KEY,
        element_justification='center',
        expand_x=True,
        expand_y=True
    )
]
_PROGRESS_TEXT_ROW = [
    sg.Text(key=PROGRESS_TEXT_KEY, text='Initialising...')
]
_PROGRESS_BAR_ROW = [
    sg.ProgressBar(key=PROGRESS_BAR_KEY, max_value=100, size=(60, 30), bar_color=PROGRESS_BAR_COLOUR)
]
_PROGRESS_OPTIONS_ROW = [
    sg.Button(key=PROGRESS_UTILITY_BUTTON_KEY, size=(13, 2), pad=((0, 5), (15, 0)), enable_events=True),
    sg.Button('Restart', key=PROGRESS_RESTART_BUTTON_KEY, size=(13, 2), pad=((5, 5), (15, 0)), enable_events=True),
    sg.Button('Exit', key=PROGRESS_EXIT_BUTTON_KEY, size=(13, 2), pad=((5, 0), (15, 0)), enable_events=True)
]
_PROGRESS_INFO_COLUMN = [
    _PROGRESS_TEXT_ROW,
    _PROGRESS_BAR_ROW,
]
_PROGRESS_OPTIONS_COLUMN = [
    _PROGRESS_OPTIONS_ROW
]
_PROGRESS_LAYOUT = [
    sg.Column(
        _PROGRESS_INFO_COLUMN,
        key=PROGRESS_COLUMN_KEY,
        expand_x=True,
        element_justification='center',
        visible=False
    ),
    sg.Column(
        _PROGRESS_OPTIONS_COLUMN,
        key=PROGRESS_OPTIONS_KEY,
        expand_x=True,
        element_justification='center',
        visible=False
    )
]
_LAYOUT = [
    _PROGRESS_LAYOUT,
    _MAIN_LAYOUT
]
_VERSION = '1.0.8'
_TITLE = f'Fixtures Excel Exporter Tool (FEET) v{_VERSION}'

WINDOW = sg.Window(
    _TITLE,
    _LAYOUT,
    size=(WINDOW_WIDTH, MAIN_WINDOW_CREATE_HEIGHT),
    resizable=False,
    icon=os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__))), 'feet.ico')
)
