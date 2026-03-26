import sys
from pathlib import Path

def error_validation(error_condition: bool, error_message: str) -> None:
    if error_condition:
        print(error_message, file=sys.stderr)
        sys.exit(1)

def error_validation_any_required_missing(args: dict) -> None:
    error_validation(
        all(value is None for value in args.values()),
        f"at least one of the following must be provided: {', '.join(args.keys())}"
    )

def error_validation_path_missing(path: Path | None) -> None:
    error_validation(
        path and not path.exists(),
        f"path '{path}' does not exist"
    )

def error_validation_file_missing(file_path: Path | None) -> None:
    error_validation(
        file_path and not file_path.is_file(),
        f"file '{file_path}' does not exist"
    )

def error_validation_directory_missing(directory_path: Path) -> None:
    error_validation(
        directory_path and not directory_path.is_dir(),
        f"directory '{directory_path}' does not exist"
    )

def error_validation_directory_is_not_extracted_pptx(directory_path: Path | None) -> None:
    error_validation(
        directory_path and not (directory_path / "[Content_Types].xml").is_file(),
        f"directory '{directory_path}' is not an extracted .pptx file"
    )

def error_validation_file_extension(file_path: Path | None, extension: str) -> None:
    error_validation(
        file_path and file_path.suffix != extension,
        f"file '{file_path}' is not a {extension} file"
    )

def error_validation_directory_exists(directory_path: Path | None) -> None:
    error_validation(
        directory_path and directory_path.is_dir(),
        f"directory '{directory_path}' already exists"
    )

def error_validation_file_exists(file_path: Path | None) -> None:
    error_validation(
        file_path and file_path.is_file(),
        f"file '{file_path}' already exists"
    )

def error_validation_slide_numbers_out_of_range(slide_numbers: list[int], slide_count: int) -> None:
    error_validation(
        any(slide_number < 1 or slide_number > slide_count for slide_number in slide_numbers),
        "invalid slide numbers"
    )

def error_validation_negative_numbers(args: dict[str, int | float | None]) -> None:
    error_validation(
        any(value is not None and value < 0 for value in args.values()),
        f"all of the following must be positive numbers: {', '.join(args.keys())}"
    )
