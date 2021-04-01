import re

tdoc_regex = re.compile(r'(?P<group>[S\d]*)-(?P<year>\d\d)(?P<tdoc_number>[\d]+)(?P<revision>r[\d][\d])?')