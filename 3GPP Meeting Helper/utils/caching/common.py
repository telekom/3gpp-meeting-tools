import hashlib
import os
import pickle
from typing import Any


def hash_file(file_path: str, chunk_size=4096) -> str|None:
    """
    Calculates the MD5 hash of a file by reading it in chunks.

    Args:
        file_path (str): The path to the file.
        chunk_size (int): The size of chunks to read from the file (in bytes).
                          Larger chunks can be faster for large files, but
                          use more memory. 4096 or 8192 are common values.

    Returns:
        str: The 32-character hexadecimal MD5 digest of the file, or None if an error occurs.
    """
    md5_hasher = hashlib.md5()
    try:
        with open(file_path, 'rb') as f:  # Open in binary read mode ('rb')
            while True:
                chunk = f.read(chunk_size)
                if not chunk:  # End of file
                    break
                md5_hasher.update(chunk)
        return md5_hasher.hexdigest()
    except FileNotFoundError:
        print(f"Error: File not found at '{file_path}'")
        return None
    except Exception as e:
        print(f"An error occurred while hashing the file: {e}")
        return None


def store_pickle_cache_for_file(
        file_path: str,
        file_prefix:str,
        data:Any,
        file_hash:str=None):

    file_folder =  os.path.dirname(file_path)
    if file_hash is None:
        file_hash = hash_file(file_path)
    target_file = os.path.join(file_folder, f'{file_prefix}_{file_hash}.pickle')
    if not os.path.exists(target_file):
        try:
            with open(target_file, 'wb') as file:
                pickle.dump(data, file)
            print(f"Object '{data}' successfully saved to '{file}'")
        except Exception as e:
            print(f"Error saving object: {e}")


def retrieve_pickle_cache_for_file(
        file_path: str,
        file_prefix:str,
        file_hash:str)->Any:

    file_folder = os.path.dirname(file_path)
    target_file = os.path.join(file_folder, f'{file_prefix}_{file_hash}.pickle')

    if not os.path.exists(target_file):
        return None

    try:
        with open(target_file, 'rb') as file:
            loaded_object = pickle.load(file)
        print(f"Object successfully loaded from '{target_file}'")
        print(f"Type of loaded object: {type(loaded_object)}")
        return loaded_object

    except FileNotFoundError:
        print(f"Error: The file '{target_file}' was not found.")
        return None
    except pickle.UnpicklingError as e:
        print(f"Error unpickling data from '{target_file}': {e}")
    except Exception as e:
        print(f"An unexpected error occurred loading {target_file}: {e}")
