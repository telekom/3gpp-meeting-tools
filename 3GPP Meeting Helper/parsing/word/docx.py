import traceback

from docx import Document


def import_agenda(agenda_file) -> None | dict[str, str]:
    """
    Imports an Agenda file
    Args:
        agenda_file (str): Path of the document to import

    Returns: AI descriptions

    """
    try:
        document = Document(agenda_file)
    except Exception as e:
        print(f'Could not open file {agenda_file}: {e}')
        return None

    all_tables = document.tables

    # First table is the topic list
    topics_table = None
    for table in all_tables:
        try:
            if table.rows[0].cells[0].text == 'AI#':
                topics_table = table
                break
        except Exception as e:
            print(f'Exception parsing topics table: {e}')

    if all_tables is None:
        print(f'Could not find topics table in {agenda_file}')
        return None

    # Table including "Lunch" string is the agenda
    agenda_table = None
    for table in all_tables[1:-1]:
        try:
            if len(table.columns) < 5:
                # Meetings are five days, so a narrower table is not valid
                continue
            lunch_cells = [cell for row in table.rows for cell in row.cells if
                           (cell.text is not None) and ('Lunch' in cell.text)]
            if len(lunch_cells) > 0:
                # Found the agenda table
                agenda_table = table
                break
        except Exception as e:
            print(f'Could not parse AIs: {e}')
            traceback.print_exc()
            pass

    # Parse AIs
    agenda_table_parsed = [(row.cells[0].text, row.cells[1].text) for row in topics_table.rows]
    agenda_table_parsed = [(entry[0].replace('\n', '').strip(), entry[1]) for entry in agenda_table_parsed if
                           (entry[0] != '') and (entry[0] != 'AI#')]
    agenda_table_parsed = [(entry[0], entry[1].split('\n')[0]) for entry in agenda_table_parsed]
    ai_descriptions = dict(agenda_table_parsed)

    print(f'{len(agenda_table_parsed)} AIs parsed: {agenda_table_parsed}')

    return ai_descriptions


