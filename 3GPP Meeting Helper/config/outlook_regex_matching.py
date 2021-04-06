# Contains regular expressions used for parsing 3GPP meeting emails

# Regular expressions used to parse email subjects. Must contain following named groups:
#   - meeting
#   - ai

# Tested with https://regex101.com/

# Examples:
# Re: [SA3#102bis-e][2.5, ki2.1][S3-210847] Evaluation of solution 2.4 MAC-S based solution
# Re: [SA3#102bis-e][2.1][KI#7][S3-211132] [5GFBS] Identifying MitM attack
# Re: [SA3#102bis-e][2.10] TR 33.851-050 --- S3-211343 approved
# [SA3#101-e][4.22] Final Status
# Re: [SA3#102bis-e][2.10.5] TR 33.851-050 --- S3-211343 approved
# [SA3#101-e][4.22] Final Status
# [SA3#102bis-e][2.7, A&A][S3-211115] UAS: Update to solution #2
# [SA3#102bis-e][1][S3 210811] E-meeting procedures
# Not supported:
# [SA3#102bis-e][2.17,sol2[S3-211091]AMF reallocation: Update to solution #2
sa3 = r'\[(?P<meeting>.*)\][\s]*\[(?P<ai>[\d]+(\.[\d]+)?(\.[\d]+)?(\.[\d]+)?)[,\s]*(?P<ai_title>[\w\s\.\&]*)?\][\s]*(?P<title>.*)'