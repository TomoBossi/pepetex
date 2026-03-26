from pathlib import Path
import xml.etree.ElementTree as ET
import tempfile

from extract import extract

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

def save_xml(tree: ET | ET.Element, output_file_path: Path) -> None:
    """
    Writes an xml.etree.ElementTree or xml.etree.ElementTree.Element object
    to a new or existing .xml file pointed at by output_file_path.
    """
    if isinstance(tree, ET.Element):
        tree = ET.ElementTree(tree)
    tree.write(output_file_path, encoding="UTF-8", xml_declaration=True)

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
    Returns the number of slides in the .pptx file pointed at by pptx_path.
    """
    if pptx_path.is_file():
        with tempfile.TemporaryDirectory() as tmp_dir:
            extract(pptx_path, tmp_dir)
            return get_slide_count_directory(tmp_dir)
    else:
        return get_slide_count_directory(pptx_path)

def get_parsed_tag(node: ET.Element) -> dict[str, str]:
    """
    Returns the parsed tag of an xml.etree.ElementTree.Element object,
    properly handling cases where the tag includes a namespace prefix.
    """
    tag = node.tag
    parsed_tag = {"tag": tag}
    if tag[0] == "{":
        split_tag = tag.split("}")
        parsed_tag["uri"] = split_tag[0][1:]
        parsed_tag["tag"] = split_tag[1]
    return parsed_tag
