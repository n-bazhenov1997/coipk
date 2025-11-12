"""This module provides utilities for working with filesystem paths using pathlib."""

from pathlib import Path

def find_files_by_masks(
    folder_path: str,
    search_masks: list[str] | str,
    max_depth: int | None = None,
    max_files: int | None = None,
) -> dict[str, Path]:
    
    """
    Search for files using masks in the specified folder path.

    :param folder_path: Path to the folder where the search starts
    :param search_masks: Set of file name masks or a single mask string
    :param max_depth: Maximum depth for folder traversal (None = no limit)
    :param max_files: Maximum number of files to be found for each mask (None = no limit)
    :return: dictionariy with pairs "filename: path"
    """

    folder_path = Path(folder_path)

    # Convert each search mask in the sheet, typle or set to a dictionary
    if isinstance(search_masks, str):
        search_masks = {search_masks: 0}
    elif (
        isinstance(search_masks, list)
        or isinstance(search_masks, tuple)
        or isinstance(search_masks, set)
    ):
        search_masks = {mask: 0 for mask in search_masks}
    else:
        raise TypeError(f'Expected list, got {type(folder_path).__name__}')

    # Create a stack of folder and add to it the inifial folder path and its depth equal to 0
    stack = [(folder_path, 0)]

    # Create a dictionary "result", in which we will write the pairs "filename: path"
    result = {}

    # Main cycle
    while stack:
        current_folder, depth = stack.pop()

        try:
            for item in current_folder.iterdir():
                if item.is_dir():
                    if max_depth is None or depth < max_depth:
                        stack.append((item, depth + 1))
                else:
                    for mask in search_masks.keys():
                       if item.math(mask):
                          pass 
        except PermissionError:
            continue

        return result
    