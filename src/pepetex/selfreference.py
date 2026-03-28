import shutil
import tempfile
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import Image

import utils
import errors
from extract import extract
from compress import compress
from namespaces import PREFIX_NAMESPACES

SELF_PIC_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld 
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:pic>
        <p:nvPicPr><p:cNvPr id="{uid}" name="" title="" />
        <p:cNvPicPr preferRelativeResize="0" /><p:nvPr /></p:nvPicPr><p:blipFill>
            <a:blip r:embed="{rel_id}">
        <a:alphaModFix /></a:blip><a:stretch><a:fillRect /></a:stretch></p:blipFill><p:spPr><a:xfrm>
            <a:off x="{x}" y="{y}" />
        <a:ext cx="{cx}" cy="{cy}" /></a:xfrm><a:prstGeom prst="rect"><a:avLst />
        </a:prstGeom><a:noFill /><a:ln><a:noFill /></a:ln></p:spPr>
    </p:pic>
</p:sld>
"""
