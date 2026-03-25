import os
import sys
from pathlib import Path

def error_validation(error_condition: bool, error_message: str) -> None:
    if error_condition:
        print(error_message, file=sys.stderr)
        sys.exit(1)

def error_validation_file_missing(file_path: str) -> None:
    error_validation(
        file_path and not os.path.exists(file_path),
        f"file '{file_path}' does not exist"
    )

def error_validation_directory_missing(directory_path: str) -> None:
    error_validation(
        directory_path and not os.path.exists(directory_path),
        f"directory '{directory_path}' does not exist"
    )

def error_validation_directory_is_not_extracted_pptx(directory_path: str) -> None:
    error_validation(
        directory_path and not os.path.exists(Path(directory_path) / "[Content_Types].xml"),
        f"directory '{directory_path}' is not an extracted .pptx file"
    )

def error_validation_file_extension(file_path, extension: str) -> None:
    error_validation(
        file_path and Path(file_path).suffix != extension,
        f"file '{file_path}' is not a {extension} file"
    )

def error_validation_directory_exists(directory_path: str) -> None:
    error_validation(
        directory_path and os.path.exists(directory_path),
        f"directory '{directory_path}' already exists"
    )

def error_validation_file_exists(file_path: str) -> None:
    error_validation(
        file_path and os.path.exists(file_path),
        f"file '{file_path}' already exists"
    )

def error_validation_slide_numbers(slide_numbers: list[int], slide_count: int) -> None:
    error_validation(
        any(slide_number < 1 or slide_number > slide_count for slide_number in slide_numbers),
        "invalid slide numbers"
    )
