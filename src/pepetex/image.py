import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import Image

import utils
import errors
from namespaces import PREFIX_NAMESPACES, DEFAULT_NAMESPACE_RELATIONSHIPS

PIC_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld 
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <p:pic>
        <p:nvPicPr>
            <p:cNvPr id="{uid}" name="{image_name}" title="{image_name}"/>
            <p:cNvPicPr preferRelativeResize="0"/>
            <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
            <a:blip r:embed="{rel_id}">
                <a:alphaModFix/>
            </a:blip>
            <a:stretch>
                <a:fillRect/>
            </a:stretch>
        </p:blipFill>
        <p:spPr>
            <a:xfrm rot="{rot}">
                <a:off x="{x}" y="{y}"/>
                <a:ext cx="{cx}" cy="{cy}"/>
            </a:xfrm>
            <a:prstGeom prst="rect">
                <a:avLst/>
            </a:prstGeom>
            <a:noFill/>
            <a:ln>
                <a:noFill/>
            </a:ln>
        </p:spPr>
    </p:pic>
</p:sld>
"""
REL_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{image_name}"/>
</Relationships>
"""

def add_image_relationship(image_name: str, slide_rels: ET.Element) -> str:
    """
    Searches for a relationship to image image_name in the slide relationships slide_rels.
    Adds the relationship to the image if it doesn't exist.
    Returns the relationship ID.
    """
    max_rel_id_number = 1
    for rel in slide_rels:
        rel_id = rel.attrib["Id"]
        if rel.attrib["Target"] == f"../media/{image_name}":
            return rel_id
        max_rel_id_number = max(max_rel_id_number, int(rel_id[3:]))
    rel_id = f"rId{max_rel_id_number + 1}"
    rel = ET.fromstring(REL_XML.format(**{"rel_id": rel_id, "image_name": image_name}))[0]
    utils.append_child_nodes(slide_rels, rel, ".")
    return rel_id

def get_image_ar(image_path: Path) -> float:
    """
    Returns the aspect ratio of the image pointed at by image_path.
    """
    with Image.open(image_path) as im:
        width, height = im.size
        return width/height

def insert_image(slide: ET.Element, image_element: ET.Element, image_index: int | None = None) -> None:
    """
    Inserts image at image_index of the slide's <p:spTree> stack.
    If image_index is None or out-of-range, it is appended at the top of the stack.
    The provided image_index is shifted by 2 so that index 0 
    corresponds to the index directly after <p:grpSpPr> in the <p:spTree> stack.
    See lines 1298-1310 at the end of page 3969 of ISO/IEC 29500-1 Third edition 2012-09-01.
    """
    sptree_xpath = f"./{{{PREFIX_NAMESPACES['p']}}}cSld/{{{PREFIX_NAMESPACES['p']}}}spTree"
    max_index = len(slide.findall(sptree_xpath)[0]) - 2
    if image_index is None or image_index > max_index:
        image_index = max_index
    utils.insert_child_nodes(slide, image_element, sptree_xpath, image_index + 2)

def image_directory(
        pptx_directory_path: Path,
        slide_numbers: list[int],
        x: int,
        y: int,
        cx: int,
        cy: int | None = None,
        rot: int = 0,
        image_path: Path | None = None,
        image_name: str | None = None,
        image_index: int | None = None
    ) -> dict:
    """
    Inserts the image pointed at by either image_path or image_name
    at index image_index of the slides indicated by slide_numbers
    of the extracted .pptx file directory pointed at by pptx_directory_path.
    If image_path is provided, image_name is ignored and the image pointed
    at by image_path is added to the pptx media directory.
    If image_name is provided instead of image_path, image_name must match
    the filename of an image file already present in the pptx media directory,
    and said file will be reused.
    Image position, size and rotation can be set with x, y, cx, cy and rot.
    If cy is not provided, it is automatically calculated based on cx
    and the aspect ratio of the image.
    Returns the name of the image in the pptx media directory,
    as well as its rId and uid per slide.
    """
    if image_path:
        image_name = utils.add_image(pptx_directory_path, image_path)
    else:
        image_path = utils.get_media_path(pptx_directory_path) / image_name
    utils.set_image_content_types(pptx_directory_path)
    slide_ids = {}
    for slide_number in slide_numbers:
        slide_rels_path = utils.get_slide_rels_path(pptx_directory_path, slide_number)
        slide_rels = ET.parse(slide_rels_path).getroot()
        slide_path = utils.get_slide_path(pptx_directory_path, slide_number)
        slide = ET.parse(slide_path).getroot()
        rel_id = add_image_relationship(image_name, slide_rels)
        uid = utils.get_new_uid(slide)
        if cy is None or cy == 0:
            cy = cx/get_image_ar(image_path)
        image_element = ET.fromstring(PIC_XML.format(**{
            "uid": uid,
            "image_name": image_name,
            "rel_id": rel_id,
            "rot": rot*60000,
            "x": x, "y": y, "cx": cx, "cy": cy
        }))[0]
        insert_image(slide, image_element, image_index)
        utils.save_xml(slide_rels, slide_rels_path, default_namespace=DEFAULT_NAMESPACE_RELATIONSHIPS)
        utils.save_xml(slide, slide_path)
        slide_ids[slide_number] = {
            "rId": rel_id,
            "uid": uid
        }
    return {
        "name": image_name,
        "slide_ids": slide_ids
    }


def image(
        pptx_path: Path,
        x: int,
        y: int,
        cx: int,
        cy: int | None = None,
        rot: int = 0,
        image_path: Path | None = None,
        image_name: str | None = None,
        image_index: int | None = None,
        slide_numbers: list[int] | None = None
    ) -> dict:
    """
    Inserts the image pointed at by either image_path or image_name
    at index image_index of the slides indicated by slide_numbers
    of the .pptx file or extracted .pptx file directory pointed at by pptx_path.
    If image_path is provided, image_name is ignored and the image pointed
    at by image_path is added to the pptx media directory.
    If image_name is provided instead of image_path, image_name must match
    the filename of an image file already present in the pptx media directory,
    and said file will be reused.
    Image position, size and rotation can be set with x, y, cx, cy and rot.
    If cy is not provided, it is automatically calculated based on cx
    and the aspect ratio of the image.
    Returns the name of the image in the pptx media directory,
    as well as its rId and uid per slide.
    """
    if slide_numbers is None:
        slide_numbers = list(range(1, 1 + utils.get_slide_count(pptx_path)))
    return utils.pptx_path_handler(pptx_path, image_directory, [slide_numbers, x, y, cx, cy, rot, image_path, image_name, image_index])

def main():
    parser = argparse.ArgumentParser(description="Inserts an image of custom postion, size and \"depth\" on specific slides.")
    parser.add_argument("-p", "--pptx-path", type=str, required=True, help="Path to a .pptx file or a directory corresponding to an extracted .pptx file.")
    parser.add_argument("-ip", "--image-path", type=str, help="Path to the image file to add to the slides. If not provided, --image-name is REQUIRED.")
    parser.add_argument("-in", "--image-name", type=str, help="Name of a preexisting image file to add to the slide. If not provided, --image-path is REQUIRED.")
    parser.add_argument("-ii", "--image-index", type=int, help="Index in the <p:spTree> stack to insert the image at. If not provided, the image is appended at the top of the stack as the foremost element of the slide. Out-of-range indexes get clamped so as to also insert the image at the top of the stack. Index 0 corresponds to directly after the <p:grpSpPr> element.")
    parser.add_argument("-s", "--slide-numbers", type=int, nargs="+", help="List of slides to modify, provided by their slide number (counting from 1). If not provided, the transition will be applied to every slide.")
    parser.add_argument("-x", "--position-x", type=int, required=True, help="Image x position, measured in English Metric Units (EMU).")
    parser.add_argument("-y", "--position-y", type=int, required=True, help="Image y position, measured in English Metric Units (EMU).")
    parser.add_argument("-cx", "--width", type=int, required=True, help="Image width, measured in English Metric Units (EMU).")
    parser.add_argument("-cy", "--height", type=int, help="Image height, measured in English Metric Units (EMU). If not provided, it is automatically set based on --width and the aspect ratio of the image.")
    parser.add_argument("-rot", "--rotation", type=int, default=0, help="Image rotation, measured in degrees. If not provided or 0, no rotation is applied to the image.")
    args = parser.parse_args()
    arg_pptx_path = Path(args.pptx_path)
    arg_image_path = None if args.image_path is None else Path(args.image_path)
    arg_image_name = args.image_name
    arg_image_index = args.image_index
    arg_slide_numbers = args.slide_numbers
    arg_x = args.position_x
    arg_y = args.position_y
    arg_cx = args.width
    arg_cy = args.height
    arg_rot = args.rotation
    errors.error_validation_path_missing(arg_pptx_path)
    errors.error_validation_any_required_missing({"-ip (--image-path)": arg_image_path, "-in (--image-name)": arg_image_name})
    errors.error_validation_slide_numbers_out_of_range(arg_slide_numbers, utils.get_slide_count(arg_pptx_path))
    errors.error_validation_negative_numbers({"-cx (--width)": arg_cx, "-cy (--height)": arg_cy, "-ii (--image_index)": arg_image_index})
    print(image(arg_pptx_path, arg_x, arg_y, arg_cx, arg_cy, arg_rot, arg_image_path, arg_image_name, arg_image_index, arg_slide_numbers))

if __name__ == "__main__":
    main()
