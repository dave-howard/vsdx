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


class GenericCell:
    """Represents a Cell element in a Layer's xml Row"""
    def __init__(self, xml: ET.Element):
        self.xml = xml

    @property
    def value(self):
        return self.xml.attrib.get('V')

    @value.setter
    def value(self, value: str):
        self.xml.attrib['V'] = str(value)

    @property
    def name(self):
        return self.xml.attrib.get('N')

    def __repr__(self):
        return f"LayerCell: name={self.name} val={self.value}"


class Layer:
    """Represents a layer in a Visio page.
    
    Layers are used to organize shapes on a page. Each layer has properties
    like name, color, visibility, and other attributes.
    """
    
    def __init__(self, xml: ET.Element, page: Page):
        """Initialize a Layer from a Row element in the Layer section.
        
        :param xml: The Row element from the Layer section
        :type xml: ET.Element
        :param page: The page this layer belongs to
        :type page: Page
        """
        self.xml = xml
        self.page = page
        self.cells = {}
        
        # Parse all cells in this row
        for cell_xml in xml.findall(f'{namespace}Cell'):
            cell = GenericCell(xml=cell_xml)
            if cell.name:
                self.cells[cell.name] = cell
    
    def __repr__(self):
        return f"<Layer id={self.id} name={self.name} visible={self.visible} color={self.color}>"
    
    @property
    def id(self) -> str:
        """Get the layer ID (IX attribute from the Row element)."""
        return self.xml.attrib.get('IX', '')
    
    @property
    def name(self) -> str:
        """Get the layer name."""
        cell = self.cells.get('Name')
        return cell.value if cell else ''
    
    @property
    def name_univ(self) -> str:
        """Get the layer universal name."""
        cell = self.cells.get('NameUniv')
        return cell.value if cell else ''
    
    @property
    def color(self) -> str:
        """Get the layer color."""
        cell = self.cells.get('Color')
        return cell.value if cell else ''
    
    @property
    def status(self) -> str:
        """Get the layer status."""
        cell = self.cells.get('Status')
        return cell.value if cell else ''
    
    @property
    def visible(self) -> bool:
        """Get whether the layer is visible."""
        cell = self.cells.get('Visible')
        return cell.value == '1' if cell else False
    
    @property
    def print(self) -> bool:
        """Get whether the layer is printable."""
        cell = self.cells.get('Print')
        return cell.value == '1' if cell else False
    
    @property
    def active(self) -> bool:
        """Get whether the layer is active."""
        cell = self.cells.get('Active')
        return cell.value == '1' if cell else False
    
    @property
    def lock(self) -> bool:
        """Get whether the layer is locked."""
        cell = self.cells.get('Lock')
        return cell.value == '1' if cell else False
    
    @property
    def snap(self) -> bool:
        """Get whether snap is enabled for the layer."""
        cell = self.cells.get('Snap')
        return cell.value == '1' if cell else False
    
    @property
    def glue(self) -> bool:
        """Get whether glue is enabled for the layer."""
        cell = self.cells.get('Glue')
        return cell.value == '1' if cell else False
    
    @property
    def color_trans(self) -> str:
        """Get the layer color transparency."""
        cell = self.cells.get('ColorTrans')
        return cell.value if cell else ''


class PagePosition(IntEnum):
    FIRST =  0
    LAST  = -1
    END   = -1
    AFTER = -2
    BEFORE= -3


class Page:
    """Represents a page or a master page in a vsdx file

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
        self.master_unique_id = None
        self.master_base_id = None
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
        pages = file_to_xml(pages_filename, self.vis.zip_file_contents)  # this contains a list of pages with rel_id and filename
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
    def is_master_page(self) -> bool:
        """Return True if this page has a master unique id and there is a match in masters xml """
        if isinstance(self.vis.masters_xml, ET.Element) and self.master_unique_id:
            master_match = f'{namespace}Master[@UniqueID="{self.master_unique_id}"]'
            master_element = self.vis.masters_xml.find(master_match)
            return master_element is not None
        return False

    @property
    def _pagesheet_xml(self):
        # get PageSheet element from pages_xml based on page_id
        ps = self.vis.pages_xml.find(f'{namespace}Page[@ID="{self.page_id}"]/{namespace}PageSheet')
        if not isinstance(ps, ET.Element):
            ps = self.vis.masters_xml.find(f'{namespace}Master[@ID="{self.page_id}"]/{namespace}PageSheet')
        return ps

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
    def layers(self) -> List[Layer]:
        """Return a list of Layer objects from the PageSheet's Layer section.
        
        Each layer is represented as a Layer object with properties like name, color, visible, etc.
        
        :return: List of Layer objects
        :rtype: List[Layer]
        """
        layers = []
        page_sheet = self._pagesheet_xml
        if page_sheet is not None:
            layer_section = page_sheet.find(f'{namespace}Section[@N="Layer"]')
            if layer_section is not None:
                for row in layer_section.findall(f'{namespace}Row'):
                    # Only create Layer if 'Name' cell has a non-empty value
                    name_cell = row.find(f'{namespace}Cell[@N="Name"]')
                    if name_cell is None or name_cell.attrib.get('V', '').strip() == '':
                        continue
                    layer = Layer(xml=row, page=self)
                    layers.append(layer)
        return layers

    def add_layer(self, name: str, color: str = '0', visible: bool = True, 
                  printable: bool = True, active: bool = False, 
                  lock: bool = False, snap: bool = True, glue: bool = True) -> Layer:
        """Add a new layer to the page.
        
        :param name: The name of the layer
        :type name: str
        :param color: The color value for the layer (default: '0')
        :type color: str
        :param visible: Whether the layer is visible (default: True)
        :type visible: bool
        :param printable: Whether the layer is printable (default: True)
        :type printable: bool
        :param active: Whether the layer is active (default: False)
        :type active: bool
        :param lock: Whether the layer is locked (default: False)
        :type lock: bool
        :param snap: Whether snap is enabled for the layer (default: True)
        :type snap: bool
        :param glue: Whether glue is enabled for the layer (default: True)
        :type glue: bool
        :return: The newly created Layer object
        :rtype: Layer
        """
        page_sheet = self._pagesheet_xml
        if page_sheet is None:
            raise ValueError("Page does not have a PageSheet")
        
        # Find or create Layer section
        layer_section = page_sheet.find(f'{namespace}Section[@N="Layer"]')
        if layer_section is None:
            layer_section = ET.fromstring(f'<Section xmlns="{namespace[1:-1]}" N="Layer"/>')
            page_sheet.append(layer_section)
        
        # Determine the next IX value
        existing_rows = layer_section.findall(f'{namespace}Row')
        next_ix = len(existing_rows)
        
        # Create new Row element
        row_xml = ET.fromstring(f'<Row xmlns="{namespace[1:-1]}" IX="{next_ix}"/>')
        
        # Add cells to the row
        cells_data = [
            ('Name', name),
            ('Color', color),
            ('Status', '0'),
            ('Visible', '1' if visible else '0'),
            ('Print', '1' if printable else '0'),
            ('Active', '1' if active else '0'),
            ('Lock', '1' if lock else '0'),
            ('Snap', '1' if snap else '0'),
            ('Glue', '1' if glue else '0'),
            ('NameUniv', name),
            ('ColorTrans', '0')
        ]
        
        for cell_name, cell_value in cells_data:
            cell_xml = ET.fromstring(f'<Cell xmlns="{namespace[1:-1]}" N="{cell_name}" V="{cell_value}"/>')
            row_xml.append(cell_xml)
        
        # Add row to layer section
        layer_section.append(row_xml)
        
        # Create and return Layer object
        return Layer(xml=row_xml, page=self)

    def get_layer_by_id(self, layer_id: str) -> Layer | None:
        """Get a layer by its ID (IX attribute).
        
        :param layer_id: The ID of the layer to find
        :type layer_id: str
        :return: The Layer object with the specified ID, or None if not found
        :rtype: Layer or None if layer doesn't exist
        """
        for layer in self.layers:
            if layer.id == layer_id:
                return layer
        return None

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

    def find_shape_by_attr(self, attr, attr_value) -> Shape:
        for s in self._shapes:
            found = s.find_shape_by_attr(attr, attr_value)
            if found:
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

    def find_shapes_by_regex(self, regex: str) -> List[Shape]:
        """Search for shapes in this page's top shape by regex"""
        return self._shapes[0].find_shapes_by_regex(regex) if len(self._shapes) else []

    @property
    def all_shapes(self):
        # return all shapes in page
        return self._shapes[0].all_shapes if len(self._shapes) else []

    def find_shape_by_property_label(self, property_label: str) -> Shape:
        """Search for shapes in this page's top shape by property label"""
        # note: use label rather than name as label is more easily visible in diagram
        return self._shapes[0].find_shape_by_property_label(property_label) if len(self._shapes) else None

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
