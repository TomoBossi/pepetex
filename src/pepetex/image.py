from pathlib import Path
import xml.etree.ElementTree as ET
import tempfile
import shutil

import argparse

from PIL import Image

from extract import extract
from compress import compress
from namespaces import namespace_uris
import utils

import errors

X_MAX = 9144000
Y_MAX = 5143500
PIC_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld 
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:pic>
        <p:nvPicPr>
            <p:cNvPr id="${uid}" name="${image_name}" title="${image_name}"/>
            <p:cNvPicPr preferRelativeResize="0"/>
            <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
            <a:blip r:embed="${rel_id}">
                <a:alphaModFix/>
            </a:blip>
            <a:stretch>
                <a:fillRect/>
            </a:stretch>
        </p:blipFill>
        <p:spPr>
            <a:xfrm rot="${rot}">
                <a:off x="${x}" y="${y}"/>
                <a:ext cx="${cx}" cy="${cy}"/>
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
<Relationship Id="${rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{image_name}"/>
"""

def add_image(pptx_directory_path: Path, image_path: Path) -> str:
    """
    Copies the image file pointed at by image_path into the media directory
    of the extracted .pptx file directory pointed at by pptx_directory_path. 
    Returns the autoincremental filename assigned to the image.
    """
    media_path = utils.get_media_path(pptx_directory_path)
    image_id = max((im.stem[5:] for im in media_path.glob("image*") if im.stem[5:].isdigit()), default=0)
    image_name = f"image{image_id}{image_path.suffix}"
    shutil.copyfile(image_path, media_path / image_name)
    return image_name

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
    rel = ET.parse(REL_XML.format({"rel_id": rel_id, "image_name": image_name})).getroot()
    utils.append_child_nodes(slide_rels, rel, ".")
    return rel_id

def get_max_uid(tree: ET.Element) -> int:
    """
    Returns the highest-numbered UID (id attribute) in the xml tree.
    """
    ids = tree.attrib.get("id", 0)
    for child in tree:
        ids.append(get_max_uid(child))
    return max(ids)

def get_new_uid(tree: ET.Element) -> int:
    """
    Returns a new available UID (id attribute) for use in the xml tree.
    """
    return get_max_uid(tree) + 1

def get_image_ar(image_path: Path) -> float:
    """
    Returns the aspect ratio of the image pointed at by image_path.
    """
    with Image.open(image_path) as im:
        width, height = im.size
        return width/height

def image_directory(
        pptx_directory_path: Path, 
        slide_number: int,
        x: int,
        y: int,
        cx: int,
        cy: int | None = None,
        rot: int = 0,
        image_path: Path | None = None,
        image_name: str | None = None,
        animation_path: Path | None = None
    ) -> str:
    """
    Appends the image pointed at by either image_path or image_name
    at the top of the stack of the slide indicated by slide_numbers
    of the extracted .pptx file directory pointed at by pptx_directory_path.
    If image_path is provided, image_name is ignored and the image pointed
    at by image_path is added to the pptx media directory.
    If image_name is provided instead of image_path, image_name must match
    the filename of an image file already present in the pptx media directory,
    and said file will be reused.
    Image position, size and rotation can be set with x, y, cx, cy and rot.
    If cy is not provided, it is automatically calculated based on cx
    and the aspect ratio of the image.
    Returns the name of the image in the pptx media directory.
    """
    if image_path:
        image_name = add_image(pptx_directory_path, image_path)
    else:
        image_path = utils.get_media_path(pptx_directory_path) / image_name
    slide_rels_path = utils.get_slide_rels_path(pptx_directory_path, slide_number)
    slide_rels = ET.parse(slide_rels_path).getroot()
    slide_path = utils.get_slide_path(pptx_directory_path, slide_number)
    slide = ET.parse(slide_path).getroot()
    rel_id = add_image_relationship(image_name, slide_rels)
    uid = get_new_uid(slide)
    if not cy:
        cy = cx/get_image_ar(image_path)
    image_element = ET.parse(PIC_XML.format({
        "uid": uid,
        "image_name": image_name,
        "rel_id": rel_id,
        "rot": rot,
        "x": x, "y": y, "cx": cx, "cy": cy
    })).getroot()[0]
    utils.append_child_nodes(slide, image_element, f"./{{{namespace_uris['p']}}}cSld/{{{namespace_uris['p']}}}spTree")
    utils.save_xml(slide, slide_path)
    return image_name

def image(
        pptx_path: Path, 
        slide_number: int,
        x: int,
        y: int,
        cx: int,
        cy: int | None = None,
        rot: int = 0,
        image_path: Path | None = None,
        image_name: str | None = None,
        animation_path: Path | None = None
    ) -> str:
    """
    Appends the image pointed at by either image_path or image_name
    to the top of the stack of the slide indicated by slide_numbers
    of the .pptx file or extracted .pptx file directory pointed at by pptx_path.
    If image_path is provided, image_name is ignored and the image pointed
    at by image_path is added to the pptx media directory.
    If image_name is provided instead of image_path, image_name must match
    the filename of an image file already present in the pptx media directory,
    and said file will be reused.
    Image position, size and rotation can be set with x, y, cx, cy and rot.
    If cy is not provided, it is automatically calculated based on cx
    and the aspect ratio of the image.
    Returns the name of the image in the pptx media directory.
    """
    if pptx_path.is_file():
        with tempfile.TemporaryDirectory() as tmp_dir:
            extract(pptx_path, tmp_dir)
            image_name = image_directory(tmp_dir, slide_number, x, y, cx, cy, rot, image_path, image_name, animation_path)
            compress(tmp_dir, pptx_path)
            return image_name
    else:
        return image_directory(pptx_path, slide_number, x, y, cx, cy, rot, image_path, image_name, animation_path)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Adds an image to the top of the stack on a specific slide.")
    parser.add_argument("-p", "--pptx-path", type=str, required=True, help="Path to a .pptx file or a directory corresponding to an extracted .pptx file.")
    parser.add_argument("-ip", "--image-path", type=str, help="Path to the image file to add to the slide. If not provided, --image-name is REQUIRED.")
    parser.add_argument("-in", "--image-name", type=str, help="Name of a preexisting image file to add to the slide. If not provided, --image-path is REQUIRED.")
    parser.add_argument("-s", "--slide-number", type=int, help="Slide to modify, provided by its slide number (counting from 1).")
    parser.add_argument("-x", "--position-x", type=int, required=True, help=f"Image x position (up to {X_MAX}).")
    parser.add_argument("-y", "--position-y", type=int, required=True, help=f"Image y position (up to {Y_MAX}).")
    parser.add_argument("-cx", "--width", type=int, required=True, help="Image width.")
    parser.add_argument("-cy", "--height", type=int, help="Image height. If not provided, it is automatically calculated based on --width and the aspect ratio of the image.")
    parser.add_argument("-rot", "--rotation", type=int, default=0, help="Image rotation.")
    args = parser.parse_args()
    arg_pptx_path = Path(args.pptx_path)
    arg_image_path = Path(args.image_path)
    arg_image_name = args.image_name
    arg_slide_number = args.slide_number
    arg_x = args.position_x
    arg_y = args.position_y
    arg_cx = args.width
    arg_cy = args.height
    arg_rot = args.rotation
    errors.error_validation_file_missing(arg_pptx_path)
    errors.error_validation_any_required_missing({"-ip (--image-path)": arg_image_path, "-in (--image-name)": arg_image_name})
    errors.error_validation_slide_numbers_out_of_range([arg_slide_number], utils.get_slide_count(arg_pptx_path))
    errors.error_validation_negative_numbers({"-cx (--width)": arg_cx, "-cy (--height)": arg_cy})
    image(arg_pptx_path, arg_slide_number, arg_x, arg_y, arg_cx, arg_cy, arg_rot, arg_image_path, arg_image_name)
