import PySimpleGUI as sg
import scraper
import parser

VERSION = '1.0.0'
TITLE = f'Fixtures Excel Exporter Tool (FEET) v{VERSION}'
TEMPLATE_ROW_1 = [
    sg.Text('Template Document:', size=(18, 1)),
    sg.In(key='-TEMPLATE DOCUMENT TEXT-', size=(60, 1), disabled=True, enable_events=True),
    sg.FileBrowse(file_types=(('Excel Documents', '*.xlsx'),))
]
OUTPUT_ROW_1 = [
    sg.Text('Output Folder:', size=(18, 1)),
    sg.In(key='-OUTPUT FOLDER TEXT-', size=(60, 1), disabled=True, enable_events=True),
    sg.FolderBrowse()
]
SYNC_ROW_2 = [
    sg.Text('Document to Sync:', size=(18, 1)),
    sg.In(key='-SYNC DOCUMENT TEXT-', size=(60, 1), disabled=True, enable_events=True),
    sg.FileBrowse(file_types=(('Excel Documents', '*.xlsx'),))
]
COLUMN_1 = [
    TEMPLATE_ROW_1,
    OUTPUT_ROW_1
]
COLUMN_2 = [
    SYNC_ROW_2
]
TAB_GROUP_ROW = [
    sg.TabGroup([[
        sg.Tab('Create', [[sg.Column(COLUMN_1, pad=15)]], expand_x=True, expand_y=True),
        sg.Tab('Sync', [[sg.Column(COLUMN_2, pad=30)]], expand_x=True, expand_y=True)
    ]], key='-TAB GROUP-', tab_location='center', enable_events=True)
]
PROCESS_ROW = [
    sg.Button('Process', key='-PROCESS BUTTON-', size=(10, 2), disabled=True)
]
COLUMN = [
    TAB_GROUP_ROW,
    PROCESS_ROW
]
LAYOUT = [[
    sg.Column(
        COLUMN,
        element_justification='center',
        vertical_alignment='center',
        expand_x=True,
        expand_y=True
    )
]]


def handle_window():
    """Create the window and handle its events

    """
    window = sg.Window(TITLE, LAYOUT, resizable=False)

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
            else:
                process_button.update(disabled=not process_button_enabled_1)
        elif curr_tab == 'Sync':
            # Get the text from the element on the 'Sync' tab
            sync_document_text = window['-SYNC DOCUMENT TEXT-'].get()

            # Determine whether the process button should be enabled
            if event == '-SYNC DOCUMENT TEXT-':
                if not process_button_enabled_2 and sync_document_text:
                    process_button.update(disabled=False)
                    process_button_enabled_2 = True
            else:
                process_button.update(disabled=not process_button_enabled_2)

        # Handle the process button being pressed
        if event == '-PROCESS BUTTON-':
            pass

    window.close()


if __name__ == '__main__':
    competition_html = scraper.get_html(scraper.COMPETITIONS_URL)
    grades_url = scraper.get_grades_url(competition_html)
    grades_html = scraper.get_html(grades_url)
    grade_urls = scraper.get_grade_urls(grades_html)
    grade_htmls = scraper.get_htmls_with_js(grade_urls)
    if grade_htmls is not None:
        roster = parser.create_roster(grade_htmls)
        print(roster)
