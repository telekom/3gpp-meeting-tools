import html2text
import pandas as pd
import re
import traceback

def revisions_file_to_dataframe(revisions_file, meeting_tdocs):
    try:
        with open(revisions_file, 'r') as file:
            tdoc_revisions_html = file.read()

        h = html2text.HTML2Text()
        # Ignore converting links from HTML
        h.ignore_links = True
        tdoc_revisions_text = h.handle(tdoc_revisions_html)
        tdocs = emails = re.findall(r'S2-[\d]{7}r[\d]{2}', tdoc_revisions_text)
        tdocs = list(set(tdocs))
        revision_list = [(tdoc[0:-3], tdoc[-2:]) for tdoc in tdocs]

        df = pd.DataFrame(revision_list, columns=['Tdoc', 'Revisions'])
        df["Revisions"] = df[["Revisions"]].apply(pd.to_numeric)
        # display(df)

        df_per_tdoc = df.groupby("Tdoc")
        maximums = df_per_tdoc.max()
        maximums.sort_values(by='Revisions', ascending=False, inplace=True)
        maximums = maximums.reset_index()
        maximums = maximums.set_index('Tdoc')
        # display(maximums)

        meeting_tdocs = pd.concat([meeting_tdocs, maximums], axis=1, sort=False)
        meeting_tdocs['Revisions'] = meeting_tdocs['Revisions'].fillna(0)
        meeting_tdocs['Revisions'] = meeting_tdocs['Revisions'].astype(int)

        indexed_df = df.set_index('Tdoc')
        return meeting_tdocs, indexed_df
    except:
        print('Could not parse revisions file')
        traceback.print_exc()
        return None, None