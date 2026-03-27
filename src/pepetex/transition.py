from pathlib import Path
import xml.etree.ElementTree as ET
import tempfile
import copy
import json

import argparse

from extract import extract
from compress import compress
from namespaces import namespace_uris
import utils

import errors

def set_transition_attrib_default(attrib: dict, default) -> dict:
    """
    Returns a copy of attrib with its default set to the provided value.\
    """
    attrib_copy = copy.deepcopy(attrib)
    attrib_copy["default"] = default
    return attrib_copy

TRANSITION_ATTRIBS = {
    "spd": {
        "name": "spd",
        "type": str,
        "default": None,
        "validations": lambda spd: spd in ("slow", "med", "fast")
    },
    "dur": {
        "name": "dur",
        "type": int,
        "default": None,
        "validations": lambda dur: dur > 0
    },
    "dir_orientation": {
        "name": "dir",
        "type": str,
        "default": None,
        "validations": lambda dir: dir in ("horz", "vert")
    },
    "dir_direction": {
        "name": "dir",
        "type": str,
        "default": None,
        "validations": lambda dir: dir in ("d", "l", "r", "u", "ld", "lu", "rd", "ru")
    },
    "isInverted": {
        "name": "isInverted",
        "type": int,
        "default": None,
        "validations": lambda isInverted: isInverted in (0, 1)
    },
    "spokes": {
        "name": "spokes",
        "type": int,
        "default": None,
        "validations": lambda spokes: spokes in (1, 2, 3, 4, 8)
    }
}
TRANSITIONS = {
    "airplane": {
        "xml": """
        <mc:Choice Requires="p15">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p15:prstTrans prst="airplane"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 1250)
        ]
    },
    "blinds": {
        "xml": """
        <mc:Choice Requires="p14">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p:blinds dir="{dir}"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 1600),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dir_orientation"], "vert")
        ]
    },
    "checker": {
        "xml": """
        <mc:Choice Requires="p14">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p:checker/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 2500)
        ]
    },
    "crush": {
        "xml": """
        <mc:Choice Requires="p15">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p15:prstTrans prst="crush"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 2000)
        ]
    },
    "curtains": {
        "xml": """
        <mc:Choice Requires="p15">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p15:prstTrans prst="curtains"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 6000)
        ]
    },
    "dissolve": {
        "xml": """
        <mc:Choice Requires="p14">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p:dissolve/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 1200)
        ]
    },
    "doors": {
        "xml": """
        <mc:Choice Requires="p14">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p14:doors dir="{dir}"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 1400),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dir_orientation"], "vert")
        ]
    },
    "flythrough": {
        "xml": """
        <mc:Choice Requires="p14">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p14:flythrough/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 800)
        ]
    },
    "fracture": {
        "xml": """
        <mc:Choice Requires="p15">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p15:prstTrans prst="fracture"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 2000)
        ]
    },
    "origami": {
        "xml": """
        <mc:Choice Requires="p15">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p15:prstTrans prst="origami"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 3250)
        ]
    },
    "prestige": {
        "xml": """
        <mc:Choice Requires="p15">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p15:prstTrans prst="prestige"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 2000)
        ]
    },
    "prism": {
        "xml": """
        <mc:Choice Requires="p14">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p14:prism isInverted="{isInverted}"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 1600),
            set_transition_attrib_default(TRANSITION_ATTRIBS["isInverted"], 1)
        ]
    },
    "ripple": {
        "xml": """
        <mc:Choice Requires="p14">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p14:ripple/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 1400)
        ]
    },
    "vortex": {
        "xml": """
        <mc:Choice Requires="p14">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p14:vortex dir="{dir}"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 4000),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dir_direction"], "r")
        ]
    },
    "wheel": {
        "xml": """
        <p:transition spd="{spd}">
            <p:wheel spokes="{spokes}"/>
        </p:transition>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["spokes"], 1)
        ]
    },
    "wind": {
        "xml": """
        <mc:Choice Requires="p15">
            <p:transition spd="{spd}" p14:dur="{dur}">
                <p15:prstTrans prst="wind"/>
            </p:transition>
        </mc:Choice>
        """,
        "attribs": [
            set_transition_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            set_transition_attrib_default(TRANSITION_ATTRIBS["dur"], 2000)
        ]
    }
}
TRANSITION_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld 
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
    xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
    xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main">
    {transition}
</p:sld>
"""
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

def get_transition_element(transition_tree: ET.Element) -> ET.Element:
    """
    Returns a transition element ready to be inserted into a slide.
    If the transition is defined using a Choice element, this function
    returns the Choice element wrapped inside an AlternateContent element.
    """
    transition_element = transition_tree[0]
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

def get_transition_arg_defaults(transition_name: str) -> dict:
    """
    Returns a dict that maps each attrib of a transition with its default value.
    """
    return {attrib["name"]: attrib["default"] for attrib in TRANSITIONS[transition_name]["attribs"]}

def build_transition_xml(transition_name: str, transition_attribs: dict | None) -> str:
    """
    Returns an xml string representing transition transition_name,
    configured using the attribute values provided by transition_attribs (or their default values),
    that can be directly parsed with ElementTree.fromstring().
    """
    if transition_attribs is None:
        transition_attribs = {}
    transition_definition = TRANSITIONS[transition_name]
    for attrib in transition_definition["attribs"]:
        if attrib["name"] not in transition_attribs:
            transition_attribs[attrib["name"]] = attrib["default"]
    configured_transition = transition_definition["xml"].format(**transition_attribs)
    return TRANSITION_XML.format(transition=configured_transition)

def transition_directory(pptx_directory_path: Path, transition_name: str, transition_attribs: dict | None, slide_numbers: list[int]) -> None:
    """
    Sets the transition transition_name as the animated transition of slides slide_numbers
    of the extracted .pptx file directory pointed at by pptx_directory_path.
    The transition can optionally be fully or partially configured using the transition_attribs dict.
    """
    for slide_number in slide_numbers: 
        slide_path = utils.get_slide_path(pptx_directory_path, slide_number)
        slide = ET.parse(slide_path).getroot()
        transition_tree = ET.fromstring(build_transition_xml(transition_name, transition_attribs))
        utils.remove_nodes(slide, f"./{{{namespace_uris['p']}}}transition")
        utils.remove_nodes(slide, f"./{{{namespace_uris['mc']}}}AlternateContent")
        transition_element = get_transition_element(transition_tree)
        transition_index = get_transition_index(slide)
        utils.insert_child_nodes(slide, transition_element, ".", transition_index)
        utils.save_xml(slide, slide_path)

def transition(pptx_path: Path, transition_name: str, transition_attribs: dict | None = None, slide_numbers: list[int] | None = None) -> None:
    """
    Sets the transition transition_name as the animated transition of slides slide_numbers
    of the .pptx file or extracted .pptx file directory pointed at by pptx_path.
    The transition can optionally be fully or partially configured using the transition_attribs dict.
    """
    if not slide_numbers:
        slide_numbers = list(range(1, 1 + utils.get_slide_count(pptx_path)))
    if pptx_path.is_file():
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_dir_path = Path(tmp_dir)
            extract(pptx_path, tmp_dir_path)
            transition_directory(tmp_dir_path, transition_name, transition_attribs, slide_numbers)
            compress(tmp_dir_path, pptx_path)
    else:
        transition_directory(pptx_path, transition_name, transition_attribs, slide_numbers)

def main():
    parser = argparse.ArgumentParser(description="Configures a slide transition animation on specific slides.")
    parser.add_argument("-p", "--pptx-path", type=str, required=True, help="Path to a .pptx file or a directory corresponding to an extracted .pptx file.")
    parser.add_argument("-t", "--transition-name", type=str, required=True, help=f"Name of the transition to be applied. Available transitions are {', '.join(TRANSITIONS.keys())}")
    parser.add_argument("-s", "--slide-numbers", type=int, nargs="+", help="List of slides to modify, provided by their slide number (counting from 1). If not provided, the transition will be applied to every slide.")
    parser.add_argument("-a", "--transition-attribs", type=str, help="Attribute values of the transition to be applied, provided as a JSON string mapping each attribute to its value by attribute name. If an attribute is not provided, its default value will be used. Attributes are named exactly as according to the PresentationML specs.")
    attribs = parser.parse_args()
    arg_pptx_path = Path(attribs.pptx_path)
    arg_transition_name = attribs.transition_name
    arg_slide_numbers = attribs.slide_numbers
    arg_transition_attribs = attribs.transition_attribs
    errors.error_validation_path_missing(arg_pptx_path)
    errors.error_validation_unavailable_transition(arg_transition_name, list(TRANSITIONS.keys()))
    errors.error_validation_invalid_transition_attribs_json(arg_transition_attribs)
    arg_transition_attribs = None if arg_transition_attribs is None else json.loads(arg_transition_attribs)
    errors.error_validation_extra_transition_attribs(arg_transition_attribs, TRANSITIONS[arg_transition_name]["attribs"])
    errors.error_validation_mistyped_transition_attribs(arg_transition_attribs, TRANSITIONS[arg_transition_name]["attribs"])
    errors.error_validation_invalid_transition_attribs(arg_transition_attribs, TRANSITIONS[arg_transition_name]["attribs"])
    errors.error_validation_slide_numbers_out_of_range(arg_slide_numbers, utils.get_slide_count(arg_pptx_path))
    transition(arg_pptx_path, arg_transition_name, arg_transition_attribs, arg_slide_numbers)

if __name__ == "__main__":
    main()
