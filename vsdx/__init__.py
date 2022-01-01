namespace = "{http://schemas.microsoft.com/office/visio/2012/main}"  # visio file name space
ext_prop_namespace = '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}'
vt_namespace = '{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}'
r_namespace = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
document_rels_namespace = "{http://schemas.openxmlformats.org/package/2006/relationships}"
cont_types_namespace = '{http://schemas.openxmlformats.org/package/2006/content-types}'

from .shapes import Cell
from .connectors import Connect
from .shapes import DataProperty
from .media import Media
from .pages import Page
from .pages import PagePosition
from .shapes import Shape
from .vsdxfile import VisioFile
