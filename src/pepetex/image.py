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

PIC_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld 
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:pic>
                <p:nvPicPr>
                    <p:cNvPr id="${uid}" name="${image_name}" title="${image_name}"/>
                    <p:cNvPicPr preferRelativeResize="0"/>
                    <p:nvPr/>
                </p:nvPicPr>
                <p:blipFill>
                    <a:blip r:embed="${ref_id}">
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
        </p:spTree>
    </p:cSld>
</p:sld>
"""

def add_image(pptx_directory_path, image_path: str) -> str:
    pass

def add_reference(pptx_directory_path, image_name: str, slide_number: int,) -> str:
    pass

def get_uid(pptx_directory_path: str, slide_number: int,) -> int:
    pass

def image_directory(pptx_directory_path: str, slide_number, x, y, cx: int, cy: int | None = None, rot: int = 0, image_path: str = "", image_name: str = "", animation_path: str = "") -> None:
    if image_path:
        image_name = add_image(pptx_directory_path, image_path)
    ref_id = add_reference(pptx_directory_path, image_name, slide_number)
    uid = get_uid(pptx_directory_path, slide_number)


def image(pptx_path: str, slide_number, x, y, cx: int, cy: int | None = None, rot: int = 0, image_path: str = "", image_name: str = "", animation_path: str = "") -> None:
    if Path(pptx_path).is_file():
        with tempfile.TemporaryDirectory() as tmp_dir:
            extract(pptx_path, tmp_dir)
            image_directory(tmp_dir, slide_number, x, y, cx, cy, rot, image_path, image_name)
            os.remove(pptx_path)
            compress(tmp_dir, pptx_path)
    else:
        image_directory(pptx_path, slide_number, x, y, cx, cy, rot, image_path, image_name)
