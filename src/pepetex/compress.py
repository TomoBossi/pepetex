import shutil
from pathlib import Path

import argparse

import errors

def compress(pptx_directory_path: str, output_pptx_path: str = "") -> None:
    """
    Zips the extracted .pptx file directory pointed at by pptx_directory_path
    and saves the output at output_pptx_path.
    """
    output_pptx_path = output_pptx_path or f"./{Path(pptx_directory_path).stem}.pptx"
    shutil.make_archive(output_pptx_path, "zip", pptx_directory_path)
    Path(f"{output_pptx_path}.zip").rename(Path(output_pptx_path))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Compresses an extracted .pptx back into a .pptx file.")
    parser.add_argument("-d", "--directory-path", type=str, required=True, help="Path to the directory corresponding to an extracted .pptx file.")
    parser.add_argument("-o", "--output-pptx-path", type=str, help="Destination .pptx file. If not provided, the output will be saved at ./<directory-path>.pptx.")
    args = parser.parse_args()
    errors.error_validation_file_exists(args.output_pptx_path)
    errors.error_validation_file_extension(args.output_pptx_path, ".pptx")
    errors.error_validation_directory_missing(args.pptx_directory_path)
    errors.error_validation_directory_is_not_extracted_pptx(args.pptx_directory_path)
    compress(args.pptx_directory_path, args.output_pptx_path)
