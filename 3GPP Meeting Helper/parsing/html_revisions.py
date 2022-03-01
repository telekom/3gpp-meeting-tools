import html2text
import pandas as pd
import re
import traceback

def extract_tdocs_from_html(html_content, is_draft=False):
    h = html2text.HTML2Text()
    # Ignore converting links from HTML
    h.ignore_links = True
    tdoc_revisions_text = h.handle(html_content)
    tdocs = re.findall(r'S2-[\d]{7}r[\d]{2}', tdoc_revisions_text)
    tdocs = list(set(tdocs))

    if not is_draft:
        tdoc_list = [(tdoc[0:-3], tdoc[-2:]) for tdoc in tdocs]
    else:
        tdoc_list = [(tdoc[0:-3], tdoc[-2:] + '*') for tdoc in tdocs]
    return tdoc_list

def revisions_file_to_dataframe(revisions_file, meeting_tdocs, drafts_file=None):
    try:
        with open(revisions_file, 'r') as file:
            tdoc_revisions_html = file.read()

        revision_list = extract_tdocs_from_html(tdoc_revisions_html)

        try:
            if drafts_file is not None:
                print('Parsing drafts files: {0}'.format(drafts_file))
                with open(drafts_file, 'r') as file:
                    tdoc_drafts_html = file.read()
                drafts_list = extract_tdocs_from_html(tdoc_drafts_html, is_draft=True)
                revision_list.extend(drafts_list)
        except:
            print('Could not open drafts file {0}'.format(drafts_file))
            # traceback.print_exc()

        df = pd.DataFrame(revision_list, columns=['Tdoc', 'Revisions'])
        # df["Revisions"] = df[["Revisions"]].apply(pd.to_numeric) # We now also have drafts
        # print(revision_list)

        df_per_tdoc = df.groupby("Tdoc")
        maximums = df_per_tdoc.max()
        maximums.sort_values(by='Revisions', ascending=False, inplace=True)
        maximums = maximums.reset_index()
        maximums = maximums.set_index('Tdoc')
        # display(maximums)

        meeting_tdocs = pd.concat([meeting_tdocs, maximums], axis=1, sort=False)
        meeting_tdocs['Revisions'] = meeting_tdocs['Revisions'].fillna(0)
        meeting_tdocs['Revisions'] = meeting_tdocs['Revisions'] #.astype(int)

        indexed_df = df.set_index('Tdoc')
        return meeting_tdocs, indexed_df
    except:
        print('Could not parse revisions file {0}. Drafts file: {1}'.format(revisions_file, drafts_file))
        # traceback.print_exc()
        return None, None