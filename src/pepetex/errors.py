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
        path is not None and not path.exists(),
        f"path '{path}' does not exist"
    )

def error_validation_file_missing(file_path: Path | None) -> None:
    error_validation(
        file_path is not None and not file_path.is_file(),
        f"file '{file_path}' does not exist"
    )

def error_validation_directory_missing(directory_path: Path) -> None:
    error_validation(
        directory_path is not None and not directory_path.is_dir(),
        f"directory '{directory_path}' does not exist"
    )

def error_validation_directory_is_not_extracted_pptx(directory_path: Path | None) -> None:
    error_validation(
        directory_path is not None and not (directory_path / "[Content_Types].xml").is_file(),
        f"directory '{directory_path}' is not an extracted .pptx file"
    )

def error_validation_file_extension(file_path: Path | None, extension: str) -> None:
    error_validation(
        file_path is not None and file_path.suffix != extension,
        f"file '{file_path}' is not a {extension} file"
    )

def error_validation_directory_exists(directory_path: Path | None) -> None:
    error_validation(
        directory_path is not None and directory_path.is_dir(),
        f"directory '{directory_path}' already exists"
    )

def error_validation_file_exists(file_path: Path | None) -> None:
    error_validation(
        file_path is not None and file_path.is_file(),
        f"file '{file_path}' already exists"
    )

def error_validation_slide_numbers_out_of_range(slide_numbers: list[int] | None, slide_count: int) -> None:
    error_validation(
        slide_numbers is not None and any(slide_number < 1 or slide_number > slide_count for slide_number in slide_numbers),
        "invalid slide numbers"
    )

def error_validation_negative_numbers(args: dict[str, int | float | None]) -> None:
    error_validation(
        any(value is not None and value < 0 for value in args.values()),
        f"all of the following must be positive numbers: {', '.join(args.keys())}"
    )

def error_validation_greater_than(args: dict[str, int | float | None], threshold: int | float) -> None:
    error_validation(
        any(value is not None and value > threshold for value in args.values()),
        f"all of the following must be less than {threshold}: {', '.join(args.keys())}"
    )

def error_validation_unavailable_transition(name: str, available_names: list[str]) -> None:
    error_validation(
        name not in available_names,
        f"invalid transition name, available transitions are {', '.join(available_names)}"
    )

def error_validation_unavailable_animation(name: str, available_names: list[str]) -> None:
    error_validation(
        name not in available_names,
        f"invalid animation name, available animations are {', '.join(available_names)}"
    )

def error_validation_unavailable_animation_start_condition(condition: str) -> None:
    error_validation(
        condition not in ["click", "with", "after"],
        "invalid animation start condition"
    )

def error_validation_invalid_attribs_json(json_str: str | None) -> None:
    if json_str is not None:
        json_valid = True
        try:
            json.loads(json_str)
        except json.JSONDecodeError, UnicodeDecodeError:
            json_valid = False
        error_validation(
            not json_valid,
            "attributes must be provided as a valid json string"
        )

def error_validation_extra_attribs(attribs: dict | None, attrib_definitions: list[dict]) -> None:
    if attribs is not None:
        available_attribs = [attrib["name"] for attrib in attrib_definitions] 
        extra_attribs = [attrib for attrib in attribs if attrib not in available_attribs]
        error_validation(
            len(extra_attribs) > 0,
            f"invalid attribute {', '.join(extra_attribs)}, available attributes are {', '.join(available_attribs)}"
        )

def error_validation_mistyped_attribs(attribs: dict | None, attrib_definitions: list[dict]) -> None:
    if attribs is not None:
        mistyped_attribs = []
        for attrib in attrib_definitions:
            if attrib["name"] in attribs and not isinstance(attribs[attrib["name"]], attrib["type"]):
                mistyped_attribs.append(attrib["name"])
        error_validation(
            len(mistyped_attribs) > 0,
            f"invalid attribute value type provided for {', '.join(mistyped_attribs)}"
        )

def error_validation_invalid_attribs(attribs: dict | None, attrib_definitions: list[dict]) -> None:
    if attribs is not None:
        invalid_attribs = []
        for attrib in attrib_definitions:
            if attrib["name"] in attribs and not attrib["validations"](attribs[attrib["name"]]):
                invalid_attribs.append(attrib["name"])
        error_validation(
            len(invalid_attribs) > 0,
            f"invalid attribute value provided for {', '.join(invalid_attribs)}"
        )

def error_validation_unused_uid(uid: int, used_uids: list[int]) -> None:
    error_validation(
        uid not in used_uids,
        f"invalid target element ID, available elements have IDs {', '.join(str(uid) for uid in used_uids)}"
    )
