from __future__ import annotations
from enum import IntEnum

from typing import List
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from .vsdxfile import VisioFile
import vsdx

import xml.etree.ElementTree as ET

import deprecation

from .connectors import Connect
from .shapes import Shape
# from .vsdxfile import file_to_xml  # todo: refactor this away - defined in set_name() to break circular imports

from vsdx import namespace, pretty_print_element


class PagePosition(IntEnum):
    FIRST =  0
    LAST  = -1
    END   = -1
    AFTER = -2
    BEFORE= -3


class Page:
    """Represents a page in a vsdx file

    :param vis: the VisioFile object the page belongs to
    :type vis: :class:`VisioFile`
    :param name: the name of the page
    :type name: str
    :param connects: a list of Connect objects in the page
    :type connects: List of :class:`Connect`

    """
    def __init__(self, xml: ET.ElementTree, filename: str, page_name: str, page_id: str, rel_id: str, vis: VisioFile):
        self._xml = xml
        self.filename = filename
        self._name = page_name
        self.page_id = page_id
        self.rel_id = rel_id
        self.rels_xml_filename = None
        self.rels_xml = None  # type: ET.ElementTree
        self.vis = vis
        self.max_id = 0
        # todo: add page id - from pages_xml - PageSheet[ID]

    def __repr__(self):
        return f"<Page name={self.name} file={self.filename} >"

    @property
    def connects(self):
        return self.get_connects()

    @deprecation.deprecated(deprecated_in="v0.5.0", removed_in="1.0.0", current_version=vsdx.__version__,
                            details="Use Page.name property instead")
    def set_name(self, value: str):
        from .vsdxfile import file_to_xml  # to break circular imports - is this really needed?
        pages_filename = self.vis._pages_filename()  # pages contains Page name, width, height, mapped to Id
        pages = file_to_xml(pages_filename)  # this contains a list of pages with rel_id and filename
        page = pages.getroot().find(f"{namespace}Page[{self.index_num + 1}]")
        if page:
            page.attrib['Name'] = value
            self.name = value
            self.vis.pages_xml = pages

    @property
    def name(self):
        if self._name:
            return self._name
        name = self.vis.pages_xml.find(f'{namespace}Page[{self.index_num + 1}]').attrib['Name']
        name_u = self.vis.pages_xml.find(f'{namespace}Page[{self.index_num + 1}]').attrib['NameU']
        return name_u or name or self._name  # return unicode name, or name if NameU not set

    @name.setter
    def name(self, value):
        self.vis.pages_xml.find(f'{namespace}Page[{self.index_num + 1}]').attrib['Name'] = str(value)
        self.vis.pages_xml.find(f'{namespace}Page[{self.index_num + 1}]').attrib['NameU'] = str(value)
        self._name = str(value)

    @property
    @deprecation.deprecated(deprecated_in="v0.5.0", removed_in="1.0.0", current_version=vsdx.__version__,
                            details="Use Page.name property instead")
    def page_name(self):
        return self.name

    @page_name.setter
    @deprecation.deprecated(deprecated_in="v0.5.0", removed_in="1.0.0", current_version=vsdx.__version__,
                            details="Use Page.name property instead")
    def page_name(self, value):
        self.name = value

    @property
    def _pagesheet_xml(self):
        # get page sheet based on 1-based in index
        #print(pretty_print_element(self.vis.pages_xml.find(f'{namespace}Page[{self.index_num+1}]')))
        return self.vis.pages_xml.find(f'{namespace}Page[{self.index_num+1}]/{namespace}PageSheet')

    @property
    def width(self):
        return float(self._pagesheet_xml.find(f'{namespace}Cell[@N="PageWidth"]').attrib.get('V'))

    @width.setter
    def width(self, value):
        value = float(value)
        self._pagesheet_xml.find(f'{namespace}Cell[@N="PageWidth"]').attrib['V'] = str(value)

    @property
    def height(self):
        return float(self._pagesheet_xml.find(f'{namespace}Cell[@N="PageHeight"]').attrib.get('V'))

    @height.setter
    def height(self, value):
        value = float(value)
        self._pagesheet_xml.find(f'{namespace}Cell[@N="PageHeight"]').attrib['V'] = str(value)

    @property
    def xml(self):
        return self._xml

    @xml.setter
    def xml(self, value):
        self._xml = value

    @property
    def _shapes(self):
        """Return a list of :class:`Shape` objects - for each 'Shapes'

        Note: typically returns one :class:`Shape` object which itself contains :class:`Shape` objects

        """
        return [Shape(xml=shapes, parent=self, page=self) for shapes in self.xml.findall(f"{namespace}Shapes")] or []

    @property
    @deprecation.deprecated(deprecated_in="0.5.0", removed_in="1.0.0", current_version=vsdx.__version__,
                            details="Use Page.child_shapes property to access top level shapes of a Page")
    def shapes(self):
        """Return a list of :class:`Shape` objects

        Note: typically returns one :class:`Shape` object which itself contains :class:`Shape` objects

        """
        return [Shape(xml=shapes, parent=self, page=self) for shapes in self.xml.findall(f"{namespace}Shapes")]

    @deprecation.deprecated(deprecated_in="0.5.0", removed_in="1.0.0", current_version=vsdx.__version__,
                            details="Use Page.child_shapes property to access top level shapes of a Page")
    def sub_shapes(self) -> List[Shape]:
        return self.child_shapes

    @property
    def child_shapes(self):
        """Return list of Shape objects at top level of VisioFile.Page

            :returns: list of `Shape` objects
            :rtype: List[Shape]
            """
        # note that self.shapes should always return a single shape
        if self._shapes:
            return self._shapes[0].child_shapes
        return []  # empty list if no top shapes object

    def set_max_ids(self):
        # get maximum shape id from xml in page
        for shapes in self._shapes:
            for shape in shapes.child_shapes:
                id = shape.get_max_id()
                if id > self.max_id:
                    self.max_id = id

        return self.max_id

    @property
    def index_num(self):
        # return zero-based index of this page in parent VisioFile.pages list
        return self.vis.pages.index(self) if self in self.vis.pages else None

    def add_connect(self, connect: Connect):
        connects = self.xml.find(f".//{namespace}Connects")
        if connects is None:
            connects = ET.fromstring(f"<Connects xmlns='{namespace[1:-1]}' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'/>")
            self.xml.getroot().append(connects)
            connects = self.xml.find(f".//{namespace}Connects")

        connects.append(connect.xml)

    def get_connects(self):
        elements = self.xml.findall(f".//{namespace}Connect")  # search recursively
        connects = [Connect(xml=e, page=self) for e in elements]
        return connects

    def get_connectors_between(self, shape_a_id: str='', shape_a_text: str='',
                              shape_b_id: str='', shape_b_text: str=''):
        shape_a = self.find_shape_by_id(shape_a_id) if shape_a_id else self.find_shape_by_text(shape_a_text)
        shape_b = self.find_shape_by_id(shape_b_id) if shape_b_id else self.find_shape_by_text(shape_b_text)
        connector_ids = set(a.ID for a in shape_a.connected_shapes).intersection(
            set(b.ID for b in shape_b.connected_shapes))

        connectors = set()
        for id in connector_ids:
            connectors.add(self.find_shape_by_id(id))
        return connectors

    def apply_text_context(self, context: dict):
        for s in self._shapes:
            s.apply_text_filter(context)

    def find_replace(self, old: str, new: str):
        for s in self._shapes:
            s.find_replace(old, new)

    def find_shape_by_id(self, shape_id) -> Shape:
        for s in self._shapes:
            found = s.find_shape_by_id(shape_id)
            if found:
                return found

    def _find_shapes_by_id(self, shape_id) -> List[Shape]:
        # return all shapes by ID - should only be used internally where ID is not unique (i.e. copying shapes)
        found = list()
        for s in self._shapes:
            found = s.find_shapes_by_id(shape_id)
            if found:
                return found
        return found

    def find_shapes_with_same_master(self, shape: Shape) -> List[Shape]:
        # return all shapes with master
        return [s for s in self.all_shapes if
                s.master_shape_ID == shape.master_shape_ID and s.master_page_ID == shape.master_page_ID]

    def find_shape_by_text(self, text: str) -> Shape:
        for s in self._shapes:
            found = s.find_shape_by_text(text)
            if found:
                return found

    def find_shapes_by_text(self, text: str) -> List[Shape]:
        shapes = list()
        for s in self._shapes:
            found = s.find_shapes_by_text(text)
            if found:
                shapes.extend(found)
        return shapes

    @property
    def all_shapes(self):
        # return all shapes within another shape, recursively, using same logic as find_shapes_by_text()
        return self._shapes[0].all_shapes if len(self._shapes) else []

    def find_shape_by_property_label(self, property_label: str) -> Shape:
        # return first matching shape with label
        # note: use label rather than name as label is more easily visible in diagram
        for s in self._shapes:
            found = s.find_shape_by_property_label(property_label)
            if found:
                return found

    def find_shapes_by_property_label(self, property_label: str) -> List[Shape]:
        # return all matching shapes with property label
        shapes = list()
        for s in self._shapes:
            found = s.find_shapes_by_property_label(property_label)
            if found:
                shapes.extend(found)
        return shapes

    def find_shape_by_property_label_value(self, property_label: str, property_value: str) -> Shape:
        # return first matching shape with label
        # note: use label rather than name as label is more easily visible in diagram
        for s in self._shapes:
            found = s.find_shape_by_property_label_value(property_label, property_value)
            if found:
                return found

    def find_shapes_by_property_label_value(self, property_label: str, property_value: str) -> List[Shape]:
        # return all matching shapes with property label
        shapes = list()
        for s in self._shapes:
            found = s.find_shapes_by_property_label_value(property_label, property_value)
            if found:
                shapes.extend(found)
        return shapes
