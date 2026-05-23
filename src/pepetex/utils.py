import copy
import shutil
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from collections.abc import Iterator

from extract import extract
from compress import compress
from namespaces import PREFIX_NAMESPACES, DEFAULT_NAMESPACE_CONTENT_TYPES

SLIDE_ID_LIST_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldIdLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <p:sldId id="{id}" r:id="{rel_id}"/>
</p:sldIdLst>
"""
SLIDE_REL_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{slide_number}.xml"/>
</Relationships>
"""
SLIDE_CONTENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml" PartName="/ppt/slides/slide{slide_number}.xml"/>
</Types>
"""
CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default ContentType="image/jpeg" Extension="jpeg"/>
    <Default ContentType="image/gif" Extension="gif"/>
    <Default ContentType="image/jpeg" Extension="jpg"/>
    <Default ContentType="image/png" Extension="png"/>
</Types>
"""

def pptx_path_handler(pptx_path: Path, directory_function, directory_function_parameters: list, recompress: bool = True):
    """
    Handles function calls and output file overwrites for pptx_path Path objects
    that may point to either a .pptx file or an extected .pptx file directory.
    """
    if pptx_path.is_file():
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_dir_path = Path(tmp_dir)
            extract(pptx_path, tmp_dir_path)
            output = directory_function(tmp_dir_path, *directory_function_parameters)
            if recompress:
                compress(tmp_dir_path, pptx_path)
            return output
    else:
        return directory_function(pptx_path, *directory_function_parameters)


def remove_nodes(tree: ET.Element, node_xpath: str, first_only: bool = False) -> None:
    """
    Removes all nodes matching node_xpath in tree.
    If first_only is True, only removes the first matching node.
    """
    for node in tree.findall(node_xpath):
        tree.remove(node)
        if first_only:
            break

def append_child_nodes(tree: ET.Element, node: ET.Element, parent_xpath: str, first_only: bool = False) -> None:
    """
    Appends node at the end of all parents matching parent_xpath in tree.
    If first_only is True, only appends node to the first matching parent.
    """
    for parent in tree.findall(parent_xpath):
        parent.append(node)
        if first_only:
            break

def insert_child_nodes(tree: ET.Element, node: ET.Element, parent_xpath: str, index: int = 0, first_only: bool = False) -> None:
    """
    Inserts node at a specific index of all parents matching parent_xpath in tree.
    If first_only is True, only inserts node in the first matching parent.
    """
    for parent in tree.findall(parent_xpath):
        parent.insert(index, node)
        if first_only:
            break

def save_xml(tree: ET.ElementTree | ET.Element, output_file_path: Path, default_namespace: str | None = None) -> None:
    """
    Writes an xml.etree.ElementTree or xml.etree.ElementTree.Element object
    to a new or existing .xml file pointed at by output_file_path.
    """
    if default_namespace is not None:
        ET.register_namespace("", default_namespace)
    if isinstance(tree, ET.Element):
        tree = ET.ElementTree(tree)
    tree.write(output_file_path, encoding="UTF-8", method="xml", xml_declaration=True)

def get_media_path(pptx_directory_path: Path) -> Path:
    """
    Returns a Path object pointing at the media folder
    found inside the extracted .pptx file directory pointed at by pptx_directory_path.
    """
    return pptx_directory_path / "ppt" / "media"

def get_slide_path(pptx_directory_path: Path, slide_number: int) -> Path:
    """
    Returns a Path object pointing at a specific slide definition .xml file
    found inside the extracted .pptx file directory pointed at by pptx_directory_path.
    """
    return pptx_directory_path / "ppt" / "slides" / f"slide{slide_number}.xml"

def get_slide_rels_path(pptx_directory_path: Path, slide_number: int) -> Path:
    """
    Returns a Path object pointing at a specific slide relationships definition .xml file
    found inside the extracted .pptx file directory pointed at by pptx_directory_path.
    """
    return pptx_directory_path / "ppt" / "slides" / "_rels" / f"slide{slide_number}.xml.rels"

def get_slide_count_directory(pptx_directory_path: Path) -> int:
    """
    Returns the number of slides found inside the extracted .pptx file directory 
    pointed at by pptx_directory_path.
    """
    slides_directory_path = pptx_directory_path / "ppt" / "slides"
    return len([entry for entry in slides_directory_path.iterdir() if entry.is_file()])

def get_slide_count(pptx_path: Path) -> int:
    """
    Returns the number of slides in the .pptx file or extracted directory pointed at by pptx_path.
    """
    return pptx_path_handler(pptx_path, get_slide_count_directory, [], recompress=False)

def get_uids(tree: ET.Element) -> list[int]:
    """
    Returns the list of UID (id attribute) in the xml tree.
    """
    uids = []
    uid = tree.attrib.get("id")
    if uid is not None and uid.isdigit():
        uids = [int(uid)]
    for child in tree:
        uids += get_uids(child)
    return uids

def get_new_uid(tree: ET.Element) -> int:
    """
    Returns a new available UID (id attribute) for use in the xml tree.
    """
    return max(get_uids(tree)) + 1

def get_new_sld_uid(tree: ET.Element) -> int:
    """
    Returns a new available non-master (< 2147483648) slide UID (id attribute) for use in the xml tree.
    """
    return max(uid for uid in get_uids(tree) if uid < 2147483648) + 1

def get_slide_uids_directory(pptx_directory_path: Path, slide_number: int) -> list[int]:
    """
    Returns the list of UIDs in use on the <p:spTree> element of slide slide_number
    of the extracted .pptx file directory pointed at by pptx_directory_path.
    """
    slide_path = get_slide_path(pptx_directory_path, slide_number)
    slide = ET.parse(slide_path).getroot()
    sptree_xpath = f"./{{{PREFIX_NAMESPACES['p']}}}cSld/{{{PREFIX_NAMESPACES['p']}}}spTree"
    sptree = slide.findall(sptree_xpath)[0]
    return get_uids(sptree)

def get_slide_uids(pptx_path: Path, slide_number: int) -> list[int]:
    """
    Returns the list of UIDs in use on the <p:spTree> element of slide slide_number
    of the .pptx file or extracted directory pointed at by pptx_path.
    """
    return pptx_path_handler(pptx_path, get_slide_uids_directory, [slide_number], recompress=False)

def set_image_content_types(pptx_directory_path: Path) -> None:
    """
    Adds defaults in [Content_Types].xml for all of the most common image file formats.
    """
    content_types_path = pptx_directory_path / "[Content_Types].xml"
    content_types_element = ET.parse(content_types_path).getroot()
    image_content_types = ET.fromstring(CONTENT_TYPES_XML)
    for image_content_type in image_content_types:
        if not any(child.attrib.get("Extension") == image_content_type.attrib["Extension"] for child in content_types_element):
            insert_child_nodes(content_types_element, image_content_type, ".")
    save_xml(content_types_element, content_types_path, default_namespace=DEFAULT_NAMESPACE_CONTENT_TYPES)

def renumber_slide_rel(pptx_directory_path: Path, slide_number: int, new_slide_number: int) -> None:
    """
    Renumbers the slide at the position given by slide_number to 
    new_slide_number in ppt/_rels/presentation.xml.rels
    """
    presentation_rels_path = pptx_directory_path / "ppt" / "_rels" / "presentation.xml.rels"
    presentation_rels = ET.parse(presentation_rels_path).getroot()
    for rel in presentation_rels:
        target = rel.attrib["Target"]
        if target[:12] == "slides/slide" and int(target[12:].split(".")[0]) == slide_number:
            rel.attrib["Target"] = f"slides/slide{new_slide_number}.xml"
            break
    save_xml(presentation_rels, presentation_rels_path)

def renumber_slide(pptx_directory_path: Path, slide_number: int, new_slide_number: int) -> None:
    """
    Renumbers slide ppt/slides/slide<slide_number>.xml and slide rels 
    ppt/slides/_rels/slide<slide_number>.xml.rels to
    slide ppt/slides/slide<new_slide_number>.xml and
    ppt/slides/_rels/slide<new_slide_number>.xml.rels.
    Also renames the slide in ppt/_rels/presentation.xml.rels
    """
    slide_path = get_slide_path(pptx_directory_path, slide_number)
    slide_rels_path = get_slide_rels_path(pptx_directory_path, slide_number)
    slide_path.rename(slide_path.with_name(f"slide{new_slide_number}.xml"))
    slide_rels_path.rename(slide_rels_path.with_name(f"slide{new_slide_number}.xml.rels"))
    renumber_slide_rel(pptx_directory_path, slide_number, new_slide_number)

def insert_slide_xmls(pptx_directory_path: Path, slide_xml: dict[str, str], slide_number: int) -> None:
    """
    Inserts a slide defined by a {"slide": "<slide_xml>", "slide_rels": "<slide_rels_xml>"} 
    dict by creating files slide ppt/slides/slide<slide_number>.xml and
    ppt/slides/_rels/slide<slide_number>.xml.rels
    """
    slide_path = get_slide_path(pptx_directory_path, slide_number)
    slide_rels_path = get_slide_rels_path(pptx_directory_path, slide_number)
    slide_path.write_text(slide_xml["slide"])
    slide_rels_path.write_text(slide_xml["slide_rels"])
    
def add_slide_relationship(slide_number: int, presentation_rels: ET.Element) -> str:
    """
    Adds a new relationship to slide at position slide_number.
    Returns the relationship ID.
    """
    max_rel_id_number = 1
    for rel in presentation_rels:
        rel_id = rel.attrib["Id"]
        max_rel_id_number = max(max_rel_id_number, int(rel_id[3:]))
    rel_id = f"rId{max_rel_id_number + 1}"
    rel = ET.fromstring(SLIDE_REL_XML.format(**{"rel_id": rel_id, "slide_number": slide_number}))[0]
    append_child_nodes(presentation_rels, rel, ".")
    return rel_id

def register_slide_rel(pptx_directory_path: Path, slide_number: int) -> str:
    """
    Registers relationship to the slide with the provided slide number
    in /ppt/_rels/presentation.xml.rels
    """
    presentation_rels_path = pptx_directory_path / "ppt" / "_rels" / "presentation.xml.rels"
    presentation_rels = ET.parse(presentation_rels_path).getroot()
    rel_id = add_slide_relationship(slide_number, presentation_rels)
    save_xml(presentation_rels, presentation_rels_path)
    return rel_id

def add_context_type_slide(pptx_directory_path: Path) -> None:
    """
    Adds a new slide to [Content_Types].xml
    """
    max_slide_number = 1
    content_types_path = pptx_directory_path / "[Content_Types].xml"
    content_types = ET.parse(content_types_path).getroot()
    for content in content_types:
        if get_parsed_tag(content)["tag"] == "Override":
            partname = content.attrib["PartName"]
            if partname[:17] == "/ppt/slides/slide":
                max_slide_number = max(max_slide_number, int(partname[17:].split(".")[0]))
    slide_content = ET.fromstring(SLIDE_CONTENT_XML.format(**{"slide_number": max_slide_number + 1}))[0]
    append_child_nodes(content_types, slide_content, ".")
    save_xml(content_types, content_types_path, default_namespace=DEFAULT_NAMESPACE_CONTENT_TYPES)

def register_slide(pptx_directory_path: Path, slide_number: int) -> None:
    """
    Registers slide with the provided slide number in /ppt/presentation.xml
    and /ppt/_rels/presentation.xml.rels
    """
    rel_id = register_slide_rel(pptx_directory_path, slide_number)
    presentation_path = pptx_directory_path / "ppt" / "presentation.xml"
    presentation = ET.parse(presentation_path).getroot()
    slide_id = ET.fromstring(SLIDE_ID_LIST_XML.format(**{"id": get_new_sld_uid(presentation), "rel_id": rel_id}))[0]
    sldidlst_xpath = f"./{{{PREFIX_NAMESPACES['p']}}}sldIdLst"
    insert_child_nodes(presentation, slide_id, sldidlst_xpath, slide_number - 1)
    save_xml(presentation, presentation_path)
    add_context_type_slide(pptx_directory_path)

def insert_slide(pptx_directory_path: Path, slide_xml: dict[str, str], slide_number: int = 1) -> None:
    """
    Inserts a new slide, given the XML definitions of the slide itself and its rels
    in the shape of a {"slide": "<slide_xml>", "slide_rels": "<slide_rels_xml>"} dict, 
    at a specific position given by slide number. The first slide is slide 1.
    """
    slides = get_slide_count_directory(pptx_directory_path)
    slide_number = max(1, min(slides + 1, slide_number))
    for i in range(slides, slide_number - 1, -1):
        renumber_slide(pptx_directory_path, i, i + 1)
    insert_slide_xmls(pptx_directory_path, slide_xml, slide_number)
    register_slide(pptx_directory_path, slide_number)

def insert_slides(pptx_directory_path: Path, slide_xmls: Iterator[dict[str, str]], slide_count: int, slide_number: int = 1) -> None:
    """
    Inserts a new group of slides, given the XML definitions of each slide and its rels
    in the shape of an iterator of slides {"slide": "<slide_xml>", "slide_rels": "<slide_rels_xml>"}
    dicts, at a specific starting position given by slide number. The first slide is slide 1.
    """
    slides = get_slide_count_directory(pptx_directory_path)
    slide_number = max(1, min(slides + 1, slide_number))
    for i in range(slides, slide_number - 1, -1):
        renumber_slide(pptx_directory_path, i, slide_count + i)
    for i, slide_xml in enumerate(slide_xmls):
        insert_slide_xmls(pptx_directory_path, slide_xml, slide_number + i)
        register_slide(pptx_directory_path, slide_number + i)

def add_image(pptx_directory_path: Path, image_path: Path) -> str:
    """
    Copies the image file pointed at by image_path into the media directory
    of the extracted .pptx file directory pointed at by pptx_directory_path. 
    Returns the autoincremental filename assigned to the image.
    """
    media_path = get_media_path(pptx_directory_path)
    image_id = max((int(im.stem[5:]) for im in media_path.glob("image*") if im.stem[5:].isdigit()), default=0) + 1
    image_name = f"image{image_id}{image_path.suffix}"
    shutil.copyfile(image_path, media_path / image_name)
    return image_name

def add_images(pptx_directory_path: Path, images_directory_path: Path) -> tuple[int, int]:
    """
    Copies the sorted files inside the directory pointed at by images_directory_path
    into the media directory of the extracted .pptx file directory pointed at
    by pptx_directory_path. image_directory_path is assumed to only contain images
    named n.jpg, where n is any natural number.
    Returns the autoincremental numbers assigned to the first and last added images.
    """
    media_path = get_media_path(pptx_directory_path)
    image_id = max((int(im.stem[5:]) for im in media_path.glob("image*") if im.stem[5:].isdigit()), default=0) + 1
    i = 0
    for i, image_path in enumerate(sorted(images_directory_path.iterdir(), key=lambda path: int(path.stem))):
        image_name = f"image{image_id + i}{image_path.suffix}"
        shutil.copyfile(image_path, media_path / image_name)
    return (image_id, image_id + i)

def get_parsed_tag(node: ET.Element) -> dict[str, str]:
    """
    Returns the parsed tag of an xml.etree.ElementTree.Element object,
    properly handling cases where the tag includes a namespace prefix.
    """
    tag = node.tag
    parsed_tag = {"tag": tag}
    if tag[0] == "{":
        split_tag = tag.split("}")
        parsed_tag["ns_uri"] = split_tag[0][1:]
        parsed_tag["tag"] = split_tag[1]
    return parsed_tag

def set_attrib_default(attrib: dict, default) -> dict:
    """
    Returns a copy of attrib with its default set to the provided value.\
    """
    attrib_copy = copy.deepcopy(attrib)
    attrib_copy["default"] = default
    return attrib_copy
