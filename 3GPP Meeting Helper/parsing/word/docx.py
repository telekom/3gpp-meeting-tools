import traceback

from docx import Document


def import_agenda(agenda_file):
    """
    Imports an Agenda file
    Args:
        agenda_file (str): Path of the document to import

    Returns:

    """
    try:
        document = Document(agenda_file)
    except Exception as e:
        print(f'Could not open file {agenda_file}: {e}')
        return None

    all_tables = document.tables

    # First table is the topic list
    topics_table = all_tables[0]

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
        except:
            traceback.print_exc()
            pass

    # Parse AIs
    agenda_table_parsed = [(row.cells[0].text, row.cells[1].text) for row in topics_table.rows]
    agenda_table_parsed = [(entry[0].replace('\n', '').strip(), entry[1]) for entry in agenda_table_parsed if
                           (entry[0] != '') and (entry[0] != 'AI#')]
    agenda_table_parsed = [(entry[0], entry[1].split('\n')[0]) for entry in agenda_table_parsed]
    ai_descriptions = dict(agenda_table_parsed)

    days = [cell.text.split(' ')[0] for cell in agenda_table.row_cells(0)][1:]
    hours_column = [cell.text for cell in agenda_table.column_cells(0)][1:]

    # Add room name
    last_hour = None
    repetitions = 0
    room_name = lambda x: 'Main room' if x == 0 else 'Breakout {0}'.format(x)
    for idx, current_hour in enumerate(hours_column):
        if last_hour == current_hour:
            repetitions += 1
            new_hour = '{0} ({1})'.format(current_hour, room_name(repetitions))
        else:
            repetitions = 0
            if 'Lunch' not in current_hour:
                new_hour = '{0} ({1})'.format(current_hour, room_name(repetitions))
            else:
                new_hour = current_hour
        last_hour = current_hour
        hours_column[idx] = new_hour

    return ai_descriptions


