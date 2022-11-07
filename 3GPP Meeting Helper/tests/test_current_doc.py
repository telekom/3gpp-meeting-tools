import unittest
import parsing.html.common as html_parser

class Test_test_current_doc(unittest.TestCase):
    def test_S2_1813381(self):
        html = '''<HTML><HEAD><META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=windows-1252"><TITLE>Current TD</TITLE></HEAD>
            <BODY>
            <table style="border:1px solid black;border-collapse:collapse; " TABLE BORDER=1 BGCOLOR=#FFFFFF CELLSPACING=0 WIDTH=100%>
            <TR VALIGN=TOP>
              <TD BORDERCOLOR=#000000 BGCOLOR =#99FFFF><FONT style=FONT-SIZE:20pt FACE= COLOR=#FF0000>AI: 6.11</FONT></TD>
              <TD BORDERCOLOR=#000000 BGCOLOR =#99FFFF><FONT style=FONT-SIZE:20pt FACE= COLOR=#000000> Current Document: </FONT></TD>
              <TD BORDERCOLOR=#000000 BGCOLOR =#99FFFF><FONT style=FONT-SIZE:20pt FACE= COLOR=#FF0000>S2-1813381</FONT></TD>
              <TD BORDERCOLOR=#000000 BGCOLOR =#99FFFF><FONT style=FONT-SIZE:20pt FACE= COLOR=#000000>LS OUT</FONT></TD>
              <TD BORDERCOLOR=#000000 BGCOLOR =#99FFFF><FONT style=FONT-SIZE:20pt FACE= COLOR=#000000>Reply LS on Data Analytics in SA WG2 NWDAF</FONT></TD>
              <TD BORDERCOLOR=#000000 BGCOLOR =#99FFFF><FONT style=FONT-SIZE:20pt FACE= COLOR=#000000>SA WG2 {TD REQ BY : Linghang Fan} </FONT></TD>
            </TR>
            </TD>
            </BODY>'''
        parsed = html_parser.parse_current_document(html)
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed, 'S2-1813381')

if __name__ == '__main__':
    unittest.main()
