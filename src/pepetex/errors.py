import sys
import json
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

def error_validation_unavailable_transition(transition_name: str, available_transition_names: list[str]) -> None:
    error_validation(
        transition_name not in available_transition_names,
        f"invalid transition name, available transitions are {', '.join(available_transition_names)}"
    )

def error_validation_invalid_transition_attribs_json(json_str: str | None) -> None:
    if json_str is not None:
        json_valid = True
        try:
            json.loads(json_str)
        except json.JSONDecodeError, UnicodeDecodeError:
            json_valid = False
        error_validation(
            not json_valid,
            "transition attributes must be provided as a valid json string"
        )

def error_validation_extra_transition_attribs(transition_attribs: dict | None, transition_attrib_definitions: list[dict]) -> None:
    if transition_attribs is not None:
        available_attribs = [attrib["name"] for attrib in transition_attrib_definitions] 
        extra_attribs = [attrib for attrib in transition_attribs if attrib not in available_attribs]
        error_validation(
            extra_attribs,
            f"invalid transition attribute {', '.join(extra_attribs)}, available attributes are {', '.join(available_attribs)}"
        )

def error_validation_mistyped_transition_attribs(transition_attribs: dict | None, transition_attrib_definitions: list[dict]) -> None:
    if transition_attribs is not None:
        mistyped_attribs = []
        for attrib in transition_attrib_definitions:
            if attrib["name"] in transition_attribs and not isinstance(transition_attribs[attrib["name"]], attrib["type"]):
                mistyped_attribs.append(attrib["name"])
        error_validation(
            mistyped_attribs,
            f"invalid transition attribute value type provided for {', '.join(mistyped_attribs)}"
        )

def error_validation_invalid_transition_attribs(transition_attribs: dict | None, transition_attrib_definitions: list[dict]) -> None:
    if transition_attribs is not None:
        invalid_attribs = []
        for attrib in transition_attrib_definitions:
            if attrib["name"] in transition_attribs and not attrib["validations"](transition_attribs[attrib["name"]]):
                invalid_attribs.append(attrib["name"])
        error_validation(
            invalid_attribs,
            f"invalid transition attribute value provided for {', '.join(invalid_attribs)}"
        )
