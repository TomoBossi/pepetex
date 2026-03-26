import zipfile
from pathlib import Path

import argparse

import errors

def extract(pptx_path: Path, output_pptx_directory_path: Path | None = None) -> None:
    """
    Unzips the .pptx file pointed at by pptx_path
    and saves the output at output_pptx_directory_path.
    """
    output_pptx_directory_path = output_pptx_directory_path or Path(".") / pptx_path.stem
    with zipfile.ZipFile(pptx_path, "r") as z:
        z.extractall(output_pptx_directory_path)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extracts the contents of a .pptx file.")
    parser.add_argument("-p", "--pptx-path", type=str, required=True, help="Path to the .pptx file.")
    parser.add_argument("-o", "--output-pptx-directory-path", type=str, help="Destination directory. If not provided, the output will be saved at ./<pptx_path with .pptx extension removed>.")
    args = parser.parse_args()
    arg_pptx_path = Path(args.pptx_path)
    arg_output_pptx_directory_path = Path(args.output_pptx_directory_path)
    errors.error_validation_file_missing(arg_pptx_path)
    errors.error_validation_file_extension(arg_pptx_path, ".pptx")
    errors.error_validation_directory_exists(arg_output_pptx_directory_path)
    extract(arg_pptx_path, arg_output_pptx_directory_path)
