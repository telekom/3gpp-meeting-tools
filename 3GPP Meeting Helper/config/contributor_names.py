import re

signature_synonyms_regex = {
    'Deutsche Telekom': re.compile(r'(\bdt\b)|(deutsche tele[kc]om)'),
    'KT':               re.compile(r'(\bkt\b)'),
    'Nokia':            re.compile(r'nokia'),
    'Qualcomm':         re.compile(r'qualcom[m]?'),
    'Huawei':           re.compile(r'(huawei)|(h[i]?s[i]?l[l]?icon)'),
    'T-Mobile USA':     re.compile(r't[-]?mobile'),
    'Verizon':          re.compile(r'verizon'),
    'ZTE':              re.compile(r'zte'),
    'NTT DoCoMo':       re.compile(r'(ntt)|(docomo)'),
    'Samsung':          re.compile(r'samsung'),
    'Blackberry':       re.compile(r'blackberry'),
    'Toyota':           re.compile(r'toyota'),
    'LG':               re.compile(r'lg[e]?( electronics)?'),
    'Cisco':            re.compile(r'cisco'),
    'Oppo':             re.compile(r'oppo'),
    'Interdigital':     re.compile(r'inter[ ]?digital|interdigtial|intedigital'),
    'Mediatek':         re.compile(r'mediatek'),
    'NEC':              re.compile(r'nec'),
    'Orange':           re.compile(r'orange'),
    'Continental':      re.compile(r'continental'),
    'Thales':           re.compile(r'thales'),
    'China Telecom':    re.compile(r'china telecom'),
    'TIM':              re.compile(r'(telecom italia)|(tim)'),
    'Ericsson':         re.compile(r'eric[s]?sson'),
    'Convida Wireless': re.compile(r'convida'),
    'China Mobile':     re.compile(r'china mobile'),
    'China Unicom':     re.compile(r'china unicom'),
    'Comtech':          re.compile(r'comtech'),
    'Gemalto':          re.compile(r'gemalto'),
    'Intel':            re.compile(r'intel'),
    'KDDI':             re.compile(r'kddi'),
    'KPN':              re.compile(r'kpn'),
    'Oracle':           re.compile(r'oracle'),
    'FirstNet':         re.compile(r'firstnet'),
    'Telstra':          re.compile(r'telstra'),
    'Vodafone':         re.compile(r'vodafone'),
    'Tencent':          re.compile(r'tencent'),
    'Sprint':           re.compile(r'sprint'),
    'Sony':             re.compile(r'sony'),
    'CableLabs':        re.compile(r'cablelabs'),
    'AT&T':             re.compile(r'at[&]?t'),
    'Charter':          re.compile(r'charter'),
    'Lenovo':           re.compile(r'(motorola mobility)|(lenovo)'),
    'SK Telecom':       re.compile(r'sk telecom'),
    'TNO':              re.compile(r'tno'),
    'Telefonica':       re.compile(r'telefonica'),
    'Softbank':         re.compile(r'softbank'),
    'Volkwagen':        re.compile(r'volkswagen'),
    'Vivo':             re.compile(r'vivo'),
    'Xiaomi':           re.compile(r'xiaomi'),
    'Alibaba':          re.compile(r'alibaba'),
    'Apple':            re.compile(r'apple'),
    'Broadcom':         re.compile(r'broadcom'),
    'CAICT':            re.compile(r'caict'),
    'CATR':             re.compile(r'catr'),
    'CATT':             re.compile(r'catt'),
    'CMCC':             re.compile(r'cmcc'),
    'UK HO':            re.compile(r'uk ho'),
    'Affirmed':         re.compile(r'affirmed'),
    'Expway':           re.compile(r'expway'),
    'HPE':              re.compile(r'hewlett packard'),
    'RAN WG1':          re.compile(r'ran wg1'),
    'RAN WG2':          re.compile(r'ran wg2'),
    'RAN WG3':          re.compile(r'ran wg3'),
    'SA WG1':           re.compile(r'sa wg1'),
    'SA WG2':           re.compile(r'sa wg2'),
    'SA WG3':           re.compile(r'sa wg3'),
    'SA WG4':           re.compile(r'sa wg4'),
    'SA WG5':           re.compile(r'sa wg5'),
    'SA WG6':           re.compile(r'sa wg6'),
    'TSG SA':           re.compile(r'tsg sa'),
    'TSG CT':           re.compile(r'tsg ct'),
    'TSG RAN':          re.compile(r'tsg ran'),
    'CT WG1':           re.compile(r'ct wg1'),
    'CT WG2':           re.compile(r'ct wg2'),
    'CT WG3':           re.compile(r'ct wg3'),
    'CT WG4':           re.compile(r'ct wg4'),
    'IETF':             re.compile(r'ietf'),
    'IEEE':             re.compile(r'ieee'),
    'BBF':              re.compile(r'bbf'),
    'Swisscom':         re.compile(r'swisscom'),
    'Sharp':            re.compile(r'sharp'),
    'Korea Telecom':    re.compile(r'korea telecom'),
    'Siemens':          re.compile(r'siemens'),
    'Sierra Wireless':  re.compile(r'sierra wireless'),
    'Fraunhofer HHI':   re.compile(r'fraunhofer hhi'),
    'GSMA':             re.compile(r'gsma'),
    'Broadband Forum':  re.compile(r'broadband forum'),
    'OneM2M':           re.compile(r'onem2m'),
    'Juniper Networks': re.compile(r'juniper'),
    'Spirent':          re.compile(r'spirent'),
    'Airbus':           re.compile(r'airbus'),
    'Mavenir':          re.compile(r'mavenir'),
    'Sennheiser':       re.compile(r'sennheiser'),
    'Sandvine':         re.compile(r'sandvine'),
    'Philips':          re.compile(r'philips'),
    'Google':           re.compile(r'google'),
    'Matrixx':          re.compile(r'matrixx'),
    'Comcast':          re.compile(r'comcast'),
    'Turkcell':         re.compile(r'turkcell'),
    'Airtel':           re.compile(r'airtel'),
    'Bosch':            re.compile(r'bosch'),
    'Sequans':          re.compile(r'sequans'),
    'Rakuten':          re.compile(r'rakute'),
    'Futurewei':        re.compile(r'futurewei'),
    'Microsoft':        re.compile(r'microsoft'),
    'Facebook':         re.compile(r'facebook|meta'),
    'DISH':             re.compile(r'dish'),
    'ETRI':             re.compile(r'etri'),
    'INSPUR':           re.compile(r'inspur')
}

recognized_vendor_list = [key for key, value in signature_synonyms_regex.items()]
recognized_vendor_list.append('Others')

contributor_columns = {}
for contributor in recognized_vendor_list:
    contributor_columns[contributor] =  ('Contributed by ' + contributor)

others_cosigners = set()
def reset_others():
    others_cosigners = set()

ls_regex = re.compile(r'(.*) \((.*)\)')

# Removes unwanted characters and mentions to TDocs (e.g. in LS TDocs)
source_replace_regex = re.compile(r'\(Rapporteur\)|\(\?\)|[\[\]\?\(\)]|(([\w\d]{2,3})-(\d\d)([\d]+))')

def get_matching_contributors(original_sources, others_cosigners, known_cosigners):
    # Assume '?' characters are typos, as they should not be there
    sources_clean   = source_replace_regex.sub('', original_sources).strip().lower()
    cosigners       = [item.strip() for item in sources_clean.split(',')]
    found_cosigners = [ key for key, regex in signature_synonyms_regex.items() if regex.search(sources_clean) is not None ]
    
    # Fix for cases where AT&T and CATT are double-counted
    if ('CATT' in found_cosigners and 'AT&T' in found_cosigners) and 'AT&T' not in cosigners:
        found_cosigners.remove('AT&T')
        # print(found_cosigners)
    
	# No "Others" co-signers, return result. 
    if len(cosigners) == len(found_cosigners):
        return found_cosigners

    # Detailed parsing
    for cosigner in cosigners:
        # Run this only once per known cosigner
        if cosigner not in known_cosigners:
            known_cosigners.add(cosigner)
            element_cosigners = [ (regex.match(cosigner) is not None) for key, regex in signature_synonyms_regex.items() ]
            if sum(element_cosigners) == 0:
                others_cosigners.add(cosigner)

    return found_cosigners

def get_contributor_columns():
    return [ value for key, value in contributor_columns.items() ]

def add_contributor_columns_to_tdoc_list(df):
    others_cosigners = set()
    known_cosigners  = set()

    df['Source (summary)'] = ''
    all_contributor_columns = get_contributor_columns()

    # One-line instead of for loop
    df[all_contributor_columns] = False
    
    for tdoc in df.index:
        tdoc_source = df.at[tdoc, 'Source']
        found_cosigners = get_matching_contributors(tdoc_source, others_cosigners, known_cosigners)

        # Fill in the summary contributor columns
        for cosigner in found_cosigners:
            contributor_column = contributor_columns[cosigner]
            df.at[tdoc, contributor_column] = True
        # Summary column
        if len(found_cosigners) > 0:
            df.at[tdoc, 'Source (summary)'] = ', '.join(found_cosigners)

    # others_cosigners contains all of the cosigners that could not be mapped to a source
    return others_cosigners