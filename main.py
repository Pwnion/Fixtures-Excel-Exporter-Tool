import subprocess

from threading import Thread
from window import *
from exception import RoundNotFoundException
from scraper import get_all_grade_htmls
from parser import create_roster
from export import create_excel, get_excel_date, update_excel


def _create_roster(date_string, create):
    """Create the roster

    Args:
        date_string(str): The date of the roster
        create(bool): Whether an Excel document is being created or updated

    Returns:
        Roster: The created roster

    """
    update_progress('Scraping required URLs...', 0)

    # Scrape the fixtures page
    try:
        grade_htmls = get_all_grade_htmls(date_string)
    except RoundNotFoundException as e:
        update_error(f'Could not update Excel document (data not found for {str(e)})')
        raise e
    except Exception as e:
        update_error('Could not scrape the required information from the internet (check your internet connection)')
        raise e

    # Change the progress message depending on creation or updating
    if create:
        update_msg = 'Parsing the data into an Excel document...'
    else:
        update_msg = 'Updating the Excel document...'

    update_progress(update_msg, 95)

    # Parse the data into a Roster object
    try:
        roster = create_roster(grade_htmls)
    except Exception as e:
        update_error('Could not create the roster')
        raise e

    return roster


def _create(values):
    """Create the Excel document

    Args:
        values(dict(str: str)): The window values

    """
    # Create the roster
    date_string = values[CALENDAR_KEY]
    roster = _create_roster(date_string, True)

    # Get template location and output folder location from window values
    template_location = values[TEMPLATE_DOCUMENT_KEY]
    output_folder_location = values[OUTPUT_FOLDER_KEY]

    # Create the Excel document
    try:
        create_excel(roster, template_location, output_folder_location)
    except Exception as e:
        update_error('Could not parse data into the Excel document')
        raise e

    update_progress('Done!', 100)
    toggle_progress_options()


def _update(values):
    """Update a previously created Excel document

    Args:
        values(dict(str: str)): Window values

    """
    # Get the location of Excel document to update
    excel_location = values[UPDATE_DOCUMENT_KEY]

    # Create the roster
    date_string = get_excel_date(excel_location)
    roster = _create_roster(date_string, False)

    # Update the Excel document
    try:
        update_excel(roster, excel_location)
    except Exception as e:
        update_error('Could not update the Excel document')
        raise e

    update_progress('Done!', 100)
    toggle_progress_options()


def _restart_program():
    """Restart the program

    """
    os.execl(sys.executable, f'"{sys.executable}"', *sys.argv)


def _handle_window():
    """Create the window and handle its events

    """
    # Get the window elements that transcend tabs
    tab_group = WINDOW[TAB_GROUP_KEY]
    process_button = WINDOW[PROCESS_BUTTON_KEY]

    # State variables
    curr_tab = None
    process_button_enabled_1 = False
    process_button_enabled_2 = False
    driver = None

    # Event Loop
    while True:
        # Handle the window being closed
        event, values = WINDOW.read()
        if event == sg.WIN_CLOSED or event == WINDOW_EXIT_EVENT:
            # If the window is closed and a driver is still active, quit it
            if driver is not None:
                driver.quit()

            break

        # Receive errors
        if event == ERROR_EVENT:
            if driver is not None:
                driver.quit()
                driver = None

            # Update progress items with the error
            WINDOW[PROGRESS_BAR_KEY].update(bar_color=PROGRESS_BAR_ERROR_COLOUR)
            update_progress(f'Error: {values[ERROR_EVENT]}', 100)
            toggle_progress_options(error=True)
            continue

        # Receive events from a running thread
        if 'THREAD' in event:
            # Receive progress updates
            if event == THREAD_PROGRESS_EVENT:
                progress = values[THREAD_PROGRESS_EVENT]
                WINDOW[PROGRESS_TEXT_KEY].update(progress[0])
                WINDOW[PROGRESS_BAR_KEY].update(progress[1])

            # Receive a driver to quit if the thread is killed while it's active
            if event == THREAD_DRIVER_EVENT:
                driver = values[THREAD_DRIVER_EVENT]

            continue

        # Receive events from the progress window
        if 'PROGRESS' in event:
            if event == PROGRESS_UTILITY_BUTTON_KEY:
                utility_button_text = WINDOW[PROGRESS_UTILITY_BUTTON_KEY].get_text()
                if utility_button_text == EXPLORER_UTILITY:
                    if curr_tab == CREATE_TAB:
                        path = values[OUTPUT_FOLDER_KEY]
                    else:
                        end = values[UPDATE_DOCUMENT_KEY].rfind('/')
                        path = values[UPDATE_DOCUMENT_KEY][:end]

                    path = path.replace('/', '\\')
                    subprocess.Popen(f'explorer "{path}"')
                elif utility_button_text == RETRY_UTILITY:
                    WINDOW[PROGRESS_TEXT_KEY].update('Initialising...')
                    WINDOW[PROGRESS_BAR_KEY].update(0, bar_color=PROGRESS_BAR_COLOUR)
                    toggle_progress_options()
                    if curr_tab == CREATE_TAB:
                        Thread(target=_create, args=(values,), daemon=True).start()
                    elif curr_tab == UPDATE_TAB:
                        Thread(target=_update, args=(values,), daemon=True).start()
            if event == PROGRESS_RESTART_BUTTON_KEY:
                _restart_program()
            if event == PROGRESS_EXIT_BUTTON_KEY:
                break

            continue

        # Get the currently selected tab
        if event == TAB_GROUP_KEY:
            curr_tab = tab_group.get()
            if curr_tab == CREATE_TAB:
                WINDOW[TAB_GROUP_KEY].set_size(size=(None, TAB_GROUP_CREATE_HEIGHT))
                change_window_height(MAIN_WINDOW_CREATE_HEIGHT)
            elif curr_tab == UPDATE_TAB:
                WINDOW[TAB_GROUP_KEY].set_size(size=(None, TAB_GROUP_UPDATE_HEIGHT))
                change_window_height(MAIN_WINDOW_UPDATE_HEIGHT)

            # Clear the input widgets being highlighted on tab change
            WINDOW[TEMPLATE_DOCUMENT_KEY if curr_tab == CREATE_TAB else UPDATE_DOCUMENT_KEY].Widget.select_clear()

        # Handle each tab differently
        if curr_tab == CREATE_TAB:
            # Get the text from the elements on the 'Create' tab
            template_document_text = WINDOW[TEMPLATE_DOCUMENT_KEY].get()
            output_folder_text = WINDOW[OUTPUT_FOLDER_KEY].get()

            # Determine whether the process button should be enabled
            if event == TEMPLATE_DOCUMENT_KEY or event == OUTPUT_FOLDER_KEY:
                if not process_button_enabled_1 and template_document_text and output_folder_text:
                    process_button.update(disabled=False)
                    process_button_enabled_1 = True
            elif event == TAB_GROUP_KEY:
                process_button.update(disabled=not process_button_enabled_1)
            elif event == PROCESS_BUTTON_KEY:
                # Switch to the progress layout and start creating the Excel document
                switch_to_progress_layout()
                Thread(target=_create, args=(values,), daemon=True).start()
        elif curr_tab == UPDATE_TAB:
            # Get the text from the element on the 'Update' tab
            update_document_text = WINDOW[UPDATE_DOCUMENT_KEY].get()

            # Determine whether the process button should be enabled
            if event == UPDATE_DOCUMENT_KEY:
                if not process_button_enabled_2 and update_document_text:
                    process_button.update(disabled=False)
                    process_button_enabled_2 = True
            elif event == TAB_GROUP_KEY:
                process_button.update(disabled=not process_button_enabled_2)
            elif event == PROCESS_BUTTON_KEY:
                # Switch to the progress layout and start updating the Excel document
                switch_to_progress_layout()
                Thread(target=_update, args=(values,), daemon=True).start()

    WINDOW.close()


if __name__ == '__main__':
    _handle_window()
