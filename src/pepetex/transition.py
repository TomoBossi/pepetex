import json
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

import utils
import errors
from namespaces import PREFIX_NAMESPACES

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
        "validations": lambda dir: dir in ("d", "l", "r", "u")
    },
    "dir_direction_full": {
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
    "pattern": {
        "name": "pattern",
        "type": str,
        "default": None,
        "validations": lambda pattern: pattern in ("diamond", "hexagon")
    },
    "spokes": {
        "name": "spokes",
        "type": int,
        "default": None,
        "validations": lambda spokes: spokes in (1, 2, 3, 4, 8)
    }
}
TRANSITION_P = """
<p:transition spd="{{spd}}">
    <{transition}/>
</p:transition>
"""
TRANSITION_P14 = """
<mc:Choice Requires="p14">
    <p:transition spd="{{spd}}" p14:dur="{{dur}}">
        <{transition}/>
    </p:transition>
</mc:Choice>
"""
TRANSITION_P15_PRESET = """
<mc:Choice Requires="p15">
    <p:transition spd="{{spd}}" p14:dur="{{dur}}">
        <p15:prstTrans prst="{prst}"/>
    </p:transition>
</mc:Choice>
"""
TRANSITIONS = {
    "airplane": {
        "xml": TRANSITION_P15_PRESET.format(**{"prst": "airplane"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 1250)
        ]
    },
    "blinds": {
        "xml": TRANSITION_P14.format(**{"transition": "p:blinds dir=\"{dir}\""}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 1600),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dir_orientation"], "vert")
        ]
    },
    "checker": {
        "xml": TRANSITION_P14.format(**{"transition": "p:checker"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 2500)
        ]
    },
    "crush": {
        "xml": TRANSITION_P15_PRESET.format(**{"prst": "crush"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 2000)
        ]
    },
    "curtains": {
        "xml": TRANSITION_P15_PRESET.format(**{"prst": "curtains"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 6000)
        ]
    },
    "dissolve": {
        "xml": TRANSITION_P14.format(**{"transition": "p:dissolve"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 1200)
        ]
    },
    "doors": {
        "xml": TRANSITION_P14.format(**{"transition": "p14:doors dir=\"{dir}\""}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 1400),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dir_orientation"], "vert")
        ]
    },
    "drape": {
        "xml": TRANSITION_P15_PRESET.format(**{"prst": "drape"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 2000)
        ]
    },
    "flythrough": {
        "xml": TRANSITION_P14.format(**{"transition": "p14:flythrough"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 800)
        ]
    },
    "fracture": {
        "xml": TRANSITION_P15_PRESET.format(**{"prst": "fracture"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 2000)
        ]
    },
    "glitter": {
        "xml": TRANSITION_P14.format(**{"transition": "p14:glitter pattern=\"{pattern}\""}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 3900),
            utils.set_attrib_default(TRANSITION_ATTRIBS["pattern"], "hexagon")
        ]
    },
    "origami": {
        "xml": TRANSITION_P15_PRESET.format(**{"prst": "origami"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 3250)
        ]
    },
    "pagecurl": {
        "xml": TRANSITION_P15_PRESET.format(**{"prst": "pageCurlDouble"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 1250)
        ]
    },
    "prestige": {
        "xml": TRANSITION_P15_PRESET.format(**{"prst": "prestige"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 2000)
        ]
    },
    "prism": {
        "xml": TRANSITION_P14.format(**{"transition": "p14:prism dir=\"{dir}\" isInverted=\"{isInverted}\""}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 1600),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dir_direction"], "l"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["isInverted"], 1)
        ]
    },
    "ripple": {
        "xml": TRANSITION_P14.format(**{"transition": "p14:ripple"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 1400)
        ]
    },
    "vortex": {
        "xml": TRANSITION_P14.format(**{"transition": "p14:vortex dir=\"{dir}\""}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 4000),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dir_direction_full"], "r")
        ]
    },
    "wheel": {
        "xml": TRANSITION_P.format(**{"transition": "p:wheel spokes=\"{spokes}\""}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["spokes"], 1)
        ]
    },
    "wind": {
        "xml": TRANSITION_P15_PRESET.format(**{"prst": "wind"}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dur"], 2000)
        ]
    },
    "wipe": {
        "xml": TRANSITION_P.format(**{"transition": "p:wipe dir=\"{dir}\""}),
        "attribs": [
            utils.set_attrib_default(TRANSITION_ATTRIBS["spd"], "slow"),
            utils.set_attrib_default(TRANSITION_ATTRIBS["dir_direction"], "l")
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
        utils.insert_child_nodes(wrapper, transition_element, f".//{{{PREFIX_NAMESPACES['mc']}}}AlternateContent")
        transition_element = wrapper.findall(f".//{{{PREFIX_NAMESPACES['mc']}}}AlternateContent")[0]
    return transition_element

def get_transition_index(slide: ET.Element) -> int:
    """
    Returns the <p:sld> children index at which the transition 
    must be inserted in order to be compliant with PresentationML.
    See lines 1361-1367 at the end of page 3970 of ISO/IEC 29500-1 Third edition 2012-09-01.
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
    return TRANSITION_XML.format(**{"transition": configured_transition})

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
        utils.remove_nodes(slide, f"./{{{PREFIX_NAMESPACES['p']}}}transition")
        utils.remove_nodes(slide, f"./{{{PREFIX_NAMESPACES['mc']}}}AlternateContent")
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
    if slide_numbers is None:
        slide_numbers = list(range(1, 1 + utils.get_slide_count(pptx_path)))
    return utils.pptx_path_handler(pptx_path, transition_directory, [transition_name, transition_attribs, slide_numbers])

def main():
    parser = argparse.ArgumentParser(description="Configures a slide transition animation on specific slides.")
    parser.add_argument("-p", "--pptx-path", type=str, required=True, help="Path to a .pptx file or a directory corresponding to an extracted .pptx file.")
    parser.add_argument("-t", "--transition-name", type=str, required=True, help=f"Name of the transition to be applied. Available transitions are {', '.join(TRANSITIONS.keys())}")
    parser.add_argument("-s", "--slide-numbers", type=int, nargs="+", help="List of slides to modify, provided by their slide number (counting from 1). If not provided, the transition will be applied to every slide.")
    parser.add_argument("-a", "--transition-attribs", type=str, help="Attribute values of the transition to be applied, provided as a JSON string mapping each attribute to its value by attribute name. If an attribute is not provided, its default value will be used. Attributes are named exactly as according to the PresentationML specs.")
    args = parser.parse_args()
    arg_pptx_path = Path(args.pptx_path)
    arg_transition_name = args.transition_name
    arg_slide_numbers = args.slide_numbers
    arg_transition_attribs = args.transition_attribs
    errors.error_validation_path_missing(arg_pptx_path)
    errors.error_validation_unavailable_transition(arg_transition_name, list(TRANSITIONS.keys()))
    errors.error_validation_invalid_attribs_json(arg_transition_attribs)
    arg_transition_attribs = None if arg_transition_attribs is None else json.loads(arg_transition_attribs)
    errors.error_validation_extra_attribs(arg_transition_attribs, TRANSITIONS[arg_transition_name]["attribs"])
    errors.error_validation_mistyped_attribs(arg_transition_attribs, TRANSITIONS[arg_transition_name]["attribs"])
    errors.error_validation_invalid_attribs(arg_transition_attribs, TRANSITIONS[arg_transition_name]["attribs"])
    errors.error_validation_slide_numbers_out_of_range(arg_slide_numbers, utils.get_slide_count(arg_pptx_path))
    transition(arg_pptx_path, arg_transition_name, arg_transition_attribs, arg_slide_numbers)

if __name__ == "__main__":
    main()
