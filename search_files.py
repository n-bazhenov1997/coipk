"""
Module: file_search_utils

This module provides a utility function for searching files in a filesystem
by filename masks. It supports single or multiple masks, limits on the number
of files per mask, and maximum folder traversal depth. The search is safe,
ignoring folders without read permissions.

Key features:
- Search by single or multiple filename masks.
- Limit the number of files found per mask.
- Limit folder traversal depth.
- Return results as a list of pathlib.Path objects.
- Cross-platform mask matching using fnmatch.
"""

from pathlib import Path
from collections.abc import Iterable
from fnmatch import fnmatch

def find_files_by_masks(
    path: str,
    masks: str | Iterable[str],
    max_depth: int | None = None,
    max_files: int | None = None,
) -> list[Path]:

    """
    Search for files in a folder and its subfolders using filename masks.

    :param path: Root folder path to start the search.
    :param masks: Single filename mask or an iterable of masks (supports wildcards '*', '?').
    :param max_depth: Maximum folder depth to traverse (None = unlimited).
    :param max_files: Maximum number of files to find for each mask (None = unlimited).
    :return: List of pathlib.Path objects for found files. Each file appears only once.

    Notes:
    - The search uses a depth-first traversal.
    - Folders without read permissions are ignored.
    - Mask matching uses fnmatch and compares only the filename (not the full path).
    """

    path = Path(path)

    # Преобразуем маски в словарь для подсчёта найденных файлов
    if isinstance(masks, str):
        masks = {masks: 0}
    elif isinstance(masks, Iterable):
        masks = {mask: 0 for mask in masks}
    else:
        raise TypeError(f'Expected string or iterable of strings, got {type(masks).__name__}')

    stack = [(path, 0)]  # DFS: стек папок с глубиной
    result = []

    while stack and masks:  # пока есть папки и маски
        current_folder, depth = stack.pop()
        try:
            for item in current_folder.iterdir():
                if item.is_dir():
                    if max_depth is None or depth < max_depth:
                        stack.append((item, depth + 1))
                else:
                    # Найти первую подходящую маску для файла
                    mask = next((m for m in masks if fnmatch(item.name, m)), None)
                    if mask:
                        result.append(item)
                        masks[mask] += 1
                        # Удаляем маску, если достигнут лимит
                        if max_files is not None and masks[mask] >= max_files:
                            del masks[mask]
        except PermissionError:
            # Игнорируем папки без доступа
            continue

    return result
