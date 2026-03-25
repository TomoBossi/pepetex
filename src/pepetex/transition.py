from pathlib import Path
import xml.etree.ElementTree as ET
import tempfile
import os

import argparse

from extract import extract
from compress import compress
from namespaces import namespace_uris
import utils

import errors

ALTERNATE_CONTENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld 
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <mc:AlternateContent>
        <mc:Fallback>
            <p:transition spd="slow">
                <p:fade/>
            </p:transition>
        </mc:Fallback>
    </mc:AlternateContent>
</p:sld>
"""

def get_transition_element(transition_tree: ET) -> ET.Element:
    """
    Returns a transition element ready to be inserted into a slide.
    If the transition is defined using a Choice element, this function
    returns the Choice element wrapped inside an AlternateContent element.
    """
    transition_element = transition_tree.getroot()[0]
    parsed_tag = utils.get_parsed_tag(transition_element)["tag"]
    if parsed_tag == "Choice":
        wrapper = ET.fromstring(ALTERNATE_CONTENT_XML)
        utils.insert_child_nodes(wrapper, transition_element, f".//{{{namespace_uris['mc']}}}AlternateContent")
        transition_element = wrapper.findall(f".//{{{namespace_uris['mc']}}}AlternateContent")[0]
    return transition_element

def get_transition_index(slide: ET.Element) -> int:
    """
    Returns the <p:sld> children index at which the transition 
    must be inserted in order to be compliant with PresentationML.
    See the end of page 3970 of ISO/IEC 29500-1 Third edition 2012-09-01.
    """
    i = 0
    for i, child in enumerate(slide):
        if utils.get_parsed_tag(child)["tag"] in ["timing", "extLst"]:
            return i
    return i + 1

def transition_directory(pptx_directory_path: str, transition_path: str, slide_numbers: list[int]) -> None:
    """
    Sets the transition defined in the .xml file pointed at by transition_path
    as the animated transition of slides slide_numbers of the extracted .pptx file directory
    pointed at by pptx_directory_path.
    """
    for slide_number in slide_numbers: 
        slide_path = utils.get_slide_path(pptx_directory_path, slide_number)
        slide = ET.parse(slide_path).getroot()
        transition_tree = ET.parse(transition_path)
        utils.remove_nodes(slide, f"./{{{namespace_uris['p']}}}transition")
        utils.remove_nodes(slide, f"./{{{namespace_uris['mc']}}}AlternateContent")
        transition_element = get_transition_element(transition_tree)
        transition_index = get_transition_index(slide)
        utils.insert_child_nodes(slide, transition_element, ".", transition_index)
        os.remove(slide_path)
        utils.save_xml(slide, slide_path)

def transition(pptx_path: str, transition_path: str, slide_numbers: list[int] | None = None) -> None:
    """
    Sets the transition defined in the .xml file pointed at by transition_path
    as the animated transition of slides slide_numbers of .pptx file 
    or extracted .pptx file directory pointed at by pptx_path.
    """
    if not slide_numbers:
        slide_numbers = list(range(1, 1 + utils.get_slide_count(pptx_path)))
    if Path(pptx_path).is_file():
        with tempfile.TemporaryDirectory() as tmp_dir:
            extract(pptx_path, tmp_dir)
            transition_directory(tmp_dir, transition_path, slide_numbers)
            os.remove(pptx_path)
            compress(tmp_dir, pptx_path)
    else:
        transition_directory(pptx_path, transition_path, slide_numbers)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Configures a slide transition animation on specific slides.")
    parser.add_argument("-p", "--pptx-path", type=str, required=True, help="Path to a .pptx file or a directory corresponding to an extracted .pptx file.")
    parser.add_argument("-t", "--transition-path", type=str, required=True, help="Path to an .xml file that defines a transition nested directly inside the root.")
    parser.add_argument("-s", "--slide-numbers", type=int, nargs="+", help="List of slides to modify, provided by their slide number (counting from 1). If not provided, the transition will be applied to every slide.")
    args = parser.parse_args()
    errors.error_validation_file_missing(args.pptx_path)
    errors.error_validation_file_extension(args.pptx_path, ".pptx")
    errors.error_validation_file_missing(args.transition_path)
    errors.error_validation_file_extension(args.transition_path, ".xml")
    errors.error_validation_slide_numbers(args.slide_numbers, utils.get_slide_count(args.pptx_path))
    transition(args.pptx_path, args.transition_path, args.slide_numbers)
