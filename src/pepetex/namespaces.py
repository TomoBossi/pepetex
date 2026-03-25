import xml.etree.ElementTree as ET

namespace_uris = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mv": "urn:schemas-microsoft-com:mac:vml",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "pvml": "urn:schemas-microsoft-com:office:powerpoint",
    "com": "http://schemas.openxmlformats.org/drawingml/2006/compatibility",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "ahyp": "http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor"
}

uri_namespaces = {uri: ns for ns, uri in namespace_uris.items()}

for ns, uri in namespace_uris.items():
    ET.register_namespace(ns, uri)
