import os
import traceback
import zipfile


def unzip_files_in_zip_file(zip_file):
    tdoc_folder = os.path.split(zip_file)[0]
    zip_ref = zipfile.ZipFile(zip_file, 'r')
    files_in_zip = zip_ref.namelist()
    # Check if is there any file in the zip that does not exist. If not, then do not extract need_to_extract = any(
    # item == False for item in map(os.path.isfile, map(lambda x: os.path.join(tdoc_folder, x), files_in_zip)))
    # Removed check whether extracting is needed, as some people reused the same file name on different document
    # versions... Added exception catch as the file may probably be already open
    try:
        zip_ref.extractall(tdoc_folder)
    except:
        print('Could not extract files')
        traceback.print_exc()
    return [os.path.join(tdoc_folder, file) for file in files_in_zip]
