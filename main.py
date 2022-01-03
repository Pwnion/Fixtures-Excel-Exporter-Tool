import os
import sys
import subprocess

from threading import Thread
from window import *
from scraper import get_all_grade_htmls
from parser import create_roster
from export import create_excel, update_excel


def _create_roster():
    """Creates the roster

    Returns:
        Roster: The created roster

    """
    # Scrape the fixtures page
    try:
        grade_htmls = get_all_grade_htmls()
    except Exception as e:
        update_error(f'Could not scrape all required information from webpages\n\nException: {str(e)}')
        raise e

    update_progress('Parsing the data into an Excel document...', 92)

    # Parse the data into a Roster object
    try:
        roster = create_roster(grade_htmls)
    except Exception as e:
        update_error(f'Could not create roster\n\nException: {str(e)}')
        raise e

    return roster


def _create(values):
    """Creates the Excel document

    Args:
        values(dict(str: str)): The window values

    """
    # Create the roster
    roster = _create_roster()

    # Get template location and output folder location from window values
    template_location = values[TEMPLATE_DOCUMENT_KEY]
    output_folder_location = values[OUTPUT_FOLDER_KEY]

    # Create the Excel document
    try:
        create_excel(roster, template_location, output_folder_location)
    except Exception as e:
        update_error(f'Could not parse data into an Excel document\n\nException: {str(e)}')
        raise e

    update_progress('Done!', 100)
    show_progress_options()


def _update(values):
    """Update a previously created Excel document

    Args:
        values(dict(str: str)): Window values

    """
    # Create the roster
    roster = _create_roster()

    # Get the location of Excel document to update
    excel_location = values[UPDATE_DOCUMENT_KEY]

    # Update the Excel document
    update_excel(roster, excel_location)

    update_progress('Done!', 100)
    show_progress_options()


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

            WINDOW[PROGRESS_BAR_KEY].update(bar_color=('#8b0000', '#8b0000'))
            update_progress('Error!', 100)
            show_progress_options(error=True)
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
            if event == PROGRESS_EXPLORER_BUTTON_KEY:
                subprocess.Popen(f'explorer "{values[OUTPUT_FOLDER_KEY]}"')
            if event == PROGRESS_RESTART_BUTTON_KEY:
                _restart_program()
            if event == PROGRESS_EXIT_BUTTON_KEY:
                break

            continue

        # Get the currently selected tab
        if event == TAB_GROUP_KEY:
            curr_tab = tab_group.get()

        # Handle each tab differently
        if curr_tab == TAB_1:
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
        elif curr_tab == TAB_2:
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
