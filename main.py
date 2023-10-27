import os

import PySimpleGUI as sg
from spreadsheet_processor import process_spreadsheet, save_spreadsheet


def main():
    # Set the theme
    sg.theme('DarkPurple1')

    # Define the layout
    layout = [
        [sg.P(), sg.Text('Unique Cell Descriptions', font="Default 12 bold", justification="center"), sg.P()],
        [sg.Text('Selecione a planilha original:')],
        [sg.Input(), sg.FileBrowse('Procurar', file_types=(('Excel Files', '*.xlsx'),))],
        [sg.Text('Selecione onde salvar a nova planilha:')],
        [sg.Input(), sg.SaveAs('Salvar Como', file_types=(('Excel Files', '*.xlsx'),))],
        [sg.VP()],
        [sg.P(), sg.Button('Processar', s=(20, 1)), sg.Button('Sair', s=(10, 1)), sg.P()]
    ]

    # Create the window
    window = sg.Window('Unique Cell Descriptions', layout, element_padding=(10, 5))

    # Event loop
    while True:
        event, values = window.read()
        if event == 'Sair' or event == sg.WINDOW_CLOSED:
            break
        elif event == 'Processar':
            # Process the spreadsheet and save the new one
            df = process_spreadsheet(values[0])
            save_spreadsheet(df, values[1])
            sg.Popup('Planilha processada com sucesso!')
            os.startfile(values[1])

    # Close the window
    window.close()


if __name__ == "__main__":
    main()
