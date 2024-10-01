import os
import traceback

import platform

if platform.system() == 'Windows':
    print('Windows System detected. Importing win32.client')
    import win32com.client

# https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppslidelayout
ppLayoutText = 2


def get_powerpoint():
    """
    Opens a PowerPoint application instance
    Returns:
        The PowerPoint instance
    """
    if platform.system() != 'Windows':
        return None

    try:
        return win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        print(f'Could not retrieve PowerPoint instance: {e}')
        traceback.print_exc()
        return None


def new_presentation():
    """
    Creates a new PowerPoint presentation

    Returns:
        A PowerPoint presentation
    """
    try:
        ppt_instance = get_powerpoint()
        return ppt_instance.Presentations.Add()
    except Exception as e:
        print(f'Could not create new presentation: {e}')
        return None


def merge_presentations(list_of_ppts, list_of_section_labels=None, headlines_for_toc=None):
    try:
        new_slideset = new_presentation()
        slides = new_slideset.Slides
        section_properties = new_slideset.SectionProperties
        inserts = []
        for idx, slideset_path in enumerate(list_of_ppts):
            file_name = os.path.splitext(slideset_path)[0]
            insert_position = slides.Count
            print('Adding to slideset in position {0}: {1} ({2})'.format(insert_position, file_name, slideset_path))

            # Skip first slide (cover slide)
            slides.InsertFromFile(slideset_path, insert_position, SlideStart=2)

            if list_of_section_labels is not None:
                section_name = '{0}, {1}'.format(list_of_section_labels[idx], headlines_for_toc[idx])
                section_properties.AddBeforeSlide(insert_position+1, section_name)

            inserts.append((insert_position+1, section_name))

        # Create TOC
        # if headlines_for_toc is not None:
            # print('Inserts: {0}'.format(first_slide_idx_of_insert))
            # print('Headlines: {0}'.format(headlines_for_toc))
            # pptLayout = new_slideset.Slides(1).CustomLayout
            # toc_slide = slides.AddSlide(slides.Count, pptLayout)
            # headlines_for_toc
        print('Merged {0} slidesets'.format(len(inserts)))
        for idx, e in enumerate(inserts):
            print('  {0}, slide {1}: {2}'.format(idx+1, e[0], e[1]))
    except Exception as e:
        print(f'Could not merge presentations: {e}')
        return
