namespace = "{http://schemas.microsoft.com/office/visio/2012/main}"  # visio file name space
ext_prop_namespace = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'
vt_namespace = '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}'
r_namespace = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
document_rels_namespace = "{http://schemas.openxmlformats.org/package/2006/relationships}"
cont_types_namespace = '{http://schemas.openxmlformats.org/package/2006/content-types}'


import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
import xml.dom.minidom as minidom   # minidom used for prettyprint


def pretty_print_element(xml: Element) -> str:
    if type(xml) is Element:
        return minidom.parseString(ET.tostring(xml)).toprettyxml()
    elif type(xml) is ET.ElementTree:
        return minidom.parseString(ET.tostring(xml.getroot())).toprettyxml()
    else:
        return f"Not an Element. type={type(xml)}"


__version__ = "0.5.4"
from .shapes import Cell
from .connectors import Connect
from .shapes import DataProperty
from .pages import Page
from .pages import PagePosition
from .shapes import Shape
from .formulae import calc_value
from .vsdxfile import VisioFile
from .media import Media
from .geometry import Geometry, GeometryRow, GeometryCell
