import PySimpleGUI as sg

from scraper import get_all_grade_htmls
from parser import create_roster
from export import create_excel

_VERSION = '1.0.0'
_TITLE = f'Fixtures Excel Exporter Tool (FEET) v{_VERSION}'
_TEMPLATE_ROW_1 = [
    sg.Text('Template Document:', size=(18, 1)),
    sg.In(key='-TEMPLATE DOCUMENT TEXT-', size=(60, 1), disabled=True, enable_events=True),
    sg.FileBrowse(file_types=(('Excel Documents', '*.xlsx'),))
]
_OUTPUT_ROW_1 = [
    sg.Text('Output Folder:', size=(18, 1)),
    sg.In(key='-OUTPUT FOLDER TEXT-', size=(60, 1), disabled=True, enable_events=True),
    sg.FolderBrowse()
]
_SYNC_ROW_2 = [
    sg.Text('Document to Sync:', size=(18, 1)),
    sg.In(key='-SYNC DOCUMENT TEXT-', size=(60, 1), disabled=True, enable_events=True),
    sg.FileBrowse(file_types=(('Excel Documents', '*.xlsx'),))
]
_COLUMN_1 = [
    _TEMPLATE_ROW_1,
    _OUTPUT_ROW_1
]
_COLUMN_2 = [
    _SYNC_ROW_2
]
_TAB_GROUP_ROW = [
    sg.TabGroup([[
        sg.Tab('Create', [[sg.Column(_COLUMN_1, pad=15)]], expand_x=True, expand_y=True),
        sg.Tab('Sync', [[sg.Column(_COLUMN_2, pad=30)]], expand_x=True, expand_y=True)
    ]], key='-TAB GROUP-', tab_location='center', enable_events=True)
]
_PROCESS_ROW = [
    sg.Button('Process', key='-PROCESS BUTTON-', size=(10, 2), disabled=True)
]
_COLUMN = [
    _TAB_GROUP_ROW,
    _PROCESS_ROW
]
_LAYOUT = [[
    sg.Column(
        _COLUMN,
        element_justification='center',
        vertical_alignment='center',
        expand_x=True,
        expand_y=True
    )
]]


def _handle_window():
    """Create the window and handle its events

    """
    window = sg.Window(_TITLE, _LAYOUT, resizable=False)

    # Get the window elements that transcend tabs
    tab_group = window['-TAB GROUP-']
    process_button = window['-PROCESS BUTTON-']

    # State variables
    curr_tab = None
    process_button_enabled_1 = False
    process_button_enabled_2 = False

    # Event Loop
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        # Get the currently selected tab
        if event == '-TAB GROUP-':
            curr_tab = tab_group.get()

        # Handle each tab differently
        if curr_tab == 'Create':
            # Get the text from the elements on the 'Create' tab
            template_document_text = window['-TEMPLATE DOCUMENT TEXT-'].get()
            output_folder_text = window['-OUTPUT FOLDER TEXT-'].get()

            # Determine whether the process button should be enabled
            if event == '-TEMPLATE DOCUMENT TEXT-' or event == '-OUTPUT FOLDER TEXT-':
                if not process_button_enabled_1 and template_document_text and output_folder_text:
                    process_button.update(disabled=False)
                    process_button_enabled_1 = True
            elif event == '-TAB GROUP-':
                process_button.update(disabled=not process_button_enabled_1)
            elif event == '-PROCESS BUTTON-':
                # Scrape the fixtures page, parse the data and create an Excel document
                grade_htmls = get_all_grade_htmls()
                roster = create_roster(grade_htmls)
                template_location = values['-TEMPLATE DOCUMENT TEXT-']
                output_folder_location = values['-OUTPUT FOLDER TEXT-']
                create_excel(roster, template_location, output_folder_location)
        elif curr_tab == 'Sync':
            # Get the text from the element on the 'Sync' tab
            sync_document_text = window['-SYNC DOCUMENT TEXT-'].get()

            # Determine whether the process button should be enabled
            if event == '-SYNC DOCUMENT TEXT-':
                if not process_button_enabled_2 and sync_document_text:
                    process_button.update(disabled=False)
                    process_button_enabled_2 = True
            elif event == '-TAB GROUP-':
                process_button.update(disabled=not process_button_enabled_2)
            elif event == '-PROCESS BUTTON-':
                pass

    window.close()


if __name__ == '__main__':
    _handle_window()
