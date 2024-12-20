# Just use a script like this if you want to profile parts of the code
# In Visual Studio, you may need to set this file as project startup
import parsing.html.common
import parsing.html.common as html_parser
import os

import parsing.html.tdocs_by_agenda

file_name = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'tests', 'tdocs_by_agenda', '136_v2.html')
meeting = parsing.html.tdocs_by_agenda.TdocsByAgendaData(file_name, v=2)
print(meeting)