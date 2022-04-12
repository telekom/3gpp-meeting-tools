import os
import traceback

import win32com.client


def get_powerpoint():
    """
    Opens a PowerPoint application instance
    Returns:
        The PowerPoint instance
    """
    try:
        return win32com.client.Dispatch("PowerPoint.Application")
    except:
        traceback.print_exc()
        return None


def new_presentation():
    """
    Creates a new PowerPoint presentation

    Returns:
        A PowerPoint presentation
    """
    ppt_instance = get_powerpoint()
    return ppt_instance.Presentations.Add()


def merge_presentations(list_of_ppts):
    new_slideset = new_presentation()
    slides = new_slideset.Slides
    section_properties = new_slideset.SectionProperties
    for idx, slideset_path in enumerate(list_of_ppts):
        file_name = os.path.splitext(slideset_path)[0]
        insert_position = slides.Count
        print('Adding to slideset in position {0}: {1} ({2})'.format(insert_position, file_name, slideset_path))
        # section_properties.AddSection(sectionIndex=idx+1, sectionName=file_name)
        # Skip first slide (cover slide)
        slides.InsertFromFile(slideset_path, insert_position, SlideStart=2)
