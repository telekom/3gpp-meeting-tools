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
# [SA3#102bis-e][2.20, KI-Cons.authoriz] [S3-211162] Authorization of multiple consumers within a NF set
# [SA3#102bis-e][2.20, KI-Co_ns.authoriz] [S3-211162] Authorization of multiple consumers within a NF set
# Not supported:
# [SA3#102bis-e][2.17,sol2[S3-211091]AMF reallocation: Update to solution #2
sa3 = r'\[(?P<meeting>.*)\][\s]*\[(?P<ai>[\d]+(\.[\d]+)?(\.[\d]+)?(\.[\d]+)?)[,\s]*(?P<ai_title>[\w\s\.\&-_]*)?\][\s]*(?P<title>.*)'

# Re: [SA2#144E, AI#8.11, S2-2102433] Discussion of Mobility Registration Update to support NR Satellite Access
sa2 = r'.*\[SA2[ ]*#(?P<meeting>[\d]+E)[ ,]+AI[#]?(?P<ai>[\d\.]+)[ ,]+(?P<tdoc>S2-(S2-)?[\d]+)\][ ]*(?P<title>.*)'