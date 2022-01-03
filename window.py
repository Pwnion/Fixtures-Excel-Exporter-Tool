import PySimpleGUI as sg

WINDOW_WIDTH = 760
MAIN_WINDOW_HEIGHT = 200
PROGRESS_WINDOW_HEIGHT = 100
PROGRESS_WINDOW_HEIGHT_WITH_OPTIONS = 150
MAIN_COLUMN_KEY = '-MAIN COLUMN-'
TEMPLATE_DOCUMENT_KEY = '-TEMPLATE DOCUMENT TEXT-'
OUTPUT_FOLDER_KEY = '-OUTPUT FOLDER TEXT-'
UPDATE_DOCUMENT_KEY = '-UPDATE DOCUMENT TEXT-'
TAB_GROUP_KEY = '-TAB GROUP-'
PROCESS_BUTTON_KEY = '-PROCESS BUTTON-'
PROGRESS_COLUMN_KEY = '-PROGRESS COLUMN-'
PROGRESS_TEXT_KEY = '-PROGRESS TEXT-'
PROGRESS_BAR_KEY = '-PROGRESS BAR-'
PROGRESS_OPTIONS_KEY = '-PROGRESS OPTIONS COLUMN-'
PROGRESS_EXPLORER_BUTTON_KEY = '-PROGRESS EXPLORER BUTTON-'
PROGRESS_RESTART_BUTTON_KEY = '-PROGRESS RESTART BUTTON-'
PROGRESS_EXIT_BUTTON_KEY = '-PROGRESS EXIT BUTTON-'
THREAD_PROGRESS_EVENT = '-THREAD PROGRESS-'
THREAD_DRIVER_EVENT = '-THREAD DRIVER-'
ERROR_EVENT = '-ERROR-'
WINDOW_EXIT_EVENT = 'Exit'
TAB_1 = 'Create'
TAB_2 = 'Update'

_TEMPLATE_ROW_1 = [
    sg.Text('Template Document:', size=(18, 1)),
    sg.In(key=TEMPLATE_DOCUMENT_KEY, size=(60, 1), disabled=True, enable_events=True),
    sg.FileBrowse(file_types=(('Excel Documents', '*.xlsx'),))
]
_OUTPUT_ROW_1 = [
    sg.Text('Output Folder:', size=(18, 1)),
    sg.In(key=OUTPUT_FOLDER_KEY, size=(60, 1), disabled=True, enable_events=True),
    sg.FolderBrowse()
]
_SYNC_ROW_2 = [
    sg.Text('Document to Update:', size=(18, 1)),
    sg.In(key=UPDATE_DOCUMENT_KEY, size=(60, 1), disabled=True, enable_events=True),
    sg.FileBrowse(file_types=(('Excel Documents', '*.xlsx'),))
]
_MAIN_COLUMN_1 = [
    _TEMPLATE_ROW_1,
    _OUTPUT_ROW_1
]
_MAIN_COLUMN_2 = [
    _SYNC_ROW_2
]
_TAB_GROUP_ROW = [
    sg.TabGroup([[
        sg.Tab('Create', [[sg.Column(_MAIN_COLUMN_1, pad=15)]], expand_x=True, expand_y=True),
        sg.Tab('Update', [[sg.Column(_MAIN_COLUMN_2, pad=30)]], expand_x=True, expand_y=True)
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
    sg.ProgressBar(key=PROGRESS_BAR_KEY, max_value=100, size=(60, 30))
]
_PROGRESS_OPTIONS_ROW = [
    sg.Button(
        'Show in Explorer',
        key=PROGRESS_EXPLORER_BUTTON_KEY,
        size=(13, 2),
        pad=((5, 5), (15, 0)),
        enable_events=True
    ),
    sg.Button('Restart', key=PROGRESS_RESTART_BUTTON_KEY, size=(13, 2), pad=((0, 5), (15, 0)), enable_events=True),
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
_VERSION = '1.0.0'
_TITLE = f'Fixtures Excel Exporter Tool (FEET) v{_VERSION}'

WINDOW = sg.Window(_TITLE, _LAYOUT, size=(WINDOW_WIDTH, MAIN_WINDOW_HEIGHT), resizable=False)


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


def show_progress_options(error=False):
    """Shows the progress options

    Args:
        error(bool): Whether an error occurred that halted the progress bar

    """
    if error:
        WINDOW[PROGRESS_EXPLORER_BUTTON_KEY].update(disabled=True)

    progress_options_column = WINDOW[PROGRESS_OPTIONS_KEY]
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
