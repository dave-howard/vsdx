from __future__ import annotations

import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from typing import Dict
from typing import List
from typing import Optional
from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from .pages import Page

from vsdx import namespace


def to_float(val: str):
    try:
        if val is None:
            return
        return float(val)
    except ValueError:
        return 0.0


class Cell:
    def __init__(self, xml: Element, shape: Shape):
        self.xml = xml
        self.shape = shape

    @property
    def value(self):
        return self.xml.attrib.get('V')

    @value.setter
    def value(self, value: str):
        self.xml.attrib['V'] = str(value)

    @property
    def formula(self):
        return self.xml.attrib.get('F')

    @formula.setter
    def formula(self, value: str):
        self.xml.attrib['F'] = str(value)

    @property
    def name(self):
        return self.xml.attrib.get('N')

    @property
    def func(self):  # assume F stands for function, i.e. F="Width*0.5"
        return self.xml.attrib.get('F')

    def __repr__(self):
        return f"Cell: name={self.name} val={self.value} func={self.func}"


class DataProperty:
    def __init__(self, *, xml: Element, shape: Shape):
        """Represents a single Data Property item associated with a Shape object"""
        name = xml.attrib.get('N')
        # get Cell element for each property of DataProperty
        label_cell = xml.find(f'{namespace}Cell[@N="Label"]')
        value_cell = xml.find(f'{namespace}Cell[@N="Value"]')
        value = value_cell.attrib.get('V') if type(value_cell) is Element else None

        if type(label_cell) is Element:
            value_type_cell = xml.find(f'{namespace}Cell[@N="Type"]')
            prompt_cell = xml.find(f'{namespace}Cell[@N="Prompt"]')
            sort_key_cell = xml.find(f'{namespace}Cell[@N="SortKey"]')

            # get values from each Cell Element
            value_type = value_type_cell.attrib.get('V') if type(value_type_cell) is Element else None
            label = label_cell.attrib.get('V') if type(label_cell) is Element else None
            prompt = prompt_cell.attrib.get('V') if type(prompt_cell) is Element else None
            sort_key = sort_key_cell.attrib.get('V') if type(sort_key_cell) is Element else None
        else:
            # over-ridden master shape properties have no label - only a name and value
            master_prop = [p for p in shape.master_shape.data_properties.values() if p.name == name][0]  # type: DataProperty
            label = master_prop.label
            value_type = master_prop.value_type
            prompt = master_prop.prompt
            sort_key = master_prop.sort_key

        # set DataProperty properties from xml
        self.name = name
        self.value = value
        self.value_type = value_type
        self.label = label
        self.prompt = prompt
        self.sort_key = sort_key

        self.shape = shape  # reference back to Shape object
        self.xml = xml  # reference to xml used to create DataProperty


class Shape:
    """Represents a single shape, or a group shape containing other shapes
    """
    def __init__(self, xml: Element, parent: Page or Shape, page: Page):
        self.xml = xml
        self.parent = parent
        self.tag = xml.tag
        self.ID = xml.attrib.get('ID', None)
        self.master_shape_ID = xml.attrib.get('MasterShape', None)
        self.master_page_ID = xml.attrib.get('Master', None)  # i.e. '2', note: the master_page.name not list index
        if self.master_page_ID is None and isinstance(parent, Shape):  # in case of a sub_shape
            self.master_page_ID = parent.master_page_ID
        self.shape_type = xml.attrib.get('Type', None)
        self.page = page

        # get Cells in Shape
        self.cells = dict()
        for e in self.xml.findall(f"{namespace}Cell"):
            cell = Cell(xml=e, shape=self)
            self.cells[cell.name] = cell
        geometry = self.xml.find(f'{namespace}Section[@N="Geometry"]')
        if geometry is not None:
            for r in geometry.findall(f"{namespace}Row"):
                row_type = r.attrib['T']
                if row_type:
                    for e in r.findall(f"{namespace}Cell"):
                        cell = Cell(xml=e, shape=self)
                        key = f"Geometry/{row_type}/{cell.name}"
                        self.cells[key] = cell
                        #print(f"added name:['{key}']={cell} {cell.xml}")

        self._data_properties = None  # internal field to hold Shape.data_propertes, set by property

    def __repr__(self):
        return f"<Shape tag={self.tag} ID={self.ID} type={self.shape_type} text='{self.text}' >"

    def copy(self, page: Optional[Page] = None) -> Shape:
        """Copy this Shape to the specified destination Page, and return the copy.

        If the destination page is not specified, the Shape is copied to its containing Page.

        :param page: The page where the new Shape will be placed.
            If not specified, the copy will be placed in the original shape's page.
        :type page: :class:`Page` (Optional)

        :return: :class:`Shape` the new copy of shape
        """
        dst_page = page or self.page
        new_shape_xml = self.page.vis.copy_shape(self.xml, dst_page)

        # set parent: location for new shape tag to be added
        if page:
            # set parent to first page Shapes tag if destination page passed
            parent = page.shapes
        else:
            # or set parent to source shapes own parent
            parent = self.parent

        return Shape(xml=new_shape_xml, parent=parent, page=dst_page)

    @property
    def master_shape(self) -> Shape:
        """Get this shapes master

        Returns this Shape's master as a Shape object (or None)

        """
        master_page = self.page.vis.get_master_page_by_id(self.master_page_ID)
        if not master_page:
            return   # None if no master page set for this Shape
        master_shape = master_page.shapes[0].sub_shapes()[0]  # there's always a single master shape in a master page

        if self.master_shape_ID is not None:
            master_sub_shape = master_shape.find_shape_by_id(self.master_shape_ID)
            return master_sub_shape

        return master_shape

    @property
    def data_properties(self) -> Dict[str, DataProperty]:
        """
        Get data properties of the shape - which labels, names, and values
        returns a dictionary of DataProperty objects indexed by property label

        :return: Dict[str, DataProperty]
        """
        if self._data_properties:
            # return cached dict if present
            return self._data_properties

        properties = dict()
        if self.master_shape:  # start with master data properties or empty dict
            properties = self.master_shape.data_properties
        properties_xml = self.xml.find(f'{namespace}Section[@N="Property"]')
        if type(properties_xml) is Element:
            property_rows = properties_xml.findall(f'{namespace}Row')
            for prop in property_rows:
                data_prop = DataProperty(xml=prop, shape=self)
                # add properties to dict to allow fast lookup by property.label
                properties[data_prop.label] = data_prop
        self._data_properties = properties  # cache for next call
        return properties

    def cell_value(self, name: str):
        cell = self.cells.get(name)
        if cell:
            return cell.value

        if self.master_page_ID is not None:
            return self.master_shape.cell_value(name)

    def cell_formula(self, name: str):
        cell = self.cells.get(name)
        if cell:
            return cell.formula

        if self.master_page_ID is not None:
            return self.master_shape.cell_formula(name)

    def set_cell_value(self, name: str, value: str):
        cell = self.cells.get(name)
        if cell:  # only set value of existing item
            cell.value = value

        elif self.master_page_ID is not None:
            master_cell_xml = self.master_shape.xml.find(f'{namespace}Cell[@N="{name}"]')
            new_cell = ET.fromstring(ET.tostring(master_cell_xml))

            self.cells[name] = Cell(xml=new_cell, shape=self)
            self.cells[name].value = value

            self.xml.append(self.cells[name].xml)

    # LineStyle="7" FillStyle="7" TextStyle="7"
    @property
    def line_style_id(self):
        return self.xml.attrib.get('LineStyle')

    @line_style_id.setter
    def line_style_id(self, value):
        self.xml.attrib['LineStyle'] = str(value)

    @property
    def fill_style_id(self):
        return self.xml.attrib.get('FillStyle')

    @fill_style_id.setter
    def fill_style_id(self, value):
        self.xml.attrib['FillStyle'] = str(value)

    @property
    def text_style_id(self):
        return self.xml.attrib.get('TextStyle')

    @text_style_id.setter
    def text_style_id(self, value):
        self.xml.attrib['TextStyle'] = str(value)

    @property
    def line_weight(self) -> float:
        val = self.cell_value('LineWeight')
        return to_float(val)

    @line_weight.setter
    def line_weight(self, value: float or str):
        self.set_cell_value('LineWeight', str(value))

    @property
    def line_color(self) -> str:
        return self.cell_value('LineColor')

    @line_color.setter
    def line_color(self, value: str):
        self.set_cell_value('LineColor', str(value))

    @property
    def x(self):
        return to_float(self.cell_value('PinX'))

    @x.setter
    def x(self, value: float or str):
        self.set_cell_value('PinX', str(value))

    @property
    def y(self):
        return to_float(self.cell_value('PinY'))

    @y.setter
    def y(self, value: float or str):
        self.set_cell_value('PinY', str(value))

    @property
    def line_to_x(self):
        return to_float(self.cell_value('Geometry/LineTo/X'))

    @line_to_x.setter
    def line_to_x(self, value):
        self.set_cell_value('Geometry/LineTo/X', str(value))

    @property
    def line_to_y(self):
        return to_float(self.cell_value('Geometry/LineTo/Y'))

    @line_to_y.setter
    def line_to_y(self, value):
        self.set_cell_value('Geometry/LineTo/Y', str(value))

    @property
    def begin_x(self):
        return to_float(self.cell_value('BeginX'))

    @begin_x.setter
    def begin_x(self, value: float or str):
        self.set_cell_value('BeginX', str(value))

    @property
    def begin_y(self):
        return to_float(self.cell_value('BeginY'))

    @begin_y.setter
    def begin_y(self, value: float or str):
        self.set_cell_value('BeginY', str(value))

    @property
    def end_x(self):
        return to_float(self.cell_value('EndX'))

    @end_x.setter
    def end_x(self, value: float or str):
        self.set_cell_value('EndX', str(value))

    @property
    def end_y(self):
        return to_float(self.cell_value('EndY'))

    @end_y.setter
    def end_y(self, value: float or str):
        self.set_cell_value('EndY', str(value))

    def move(self, x_delta: float, y_delta: float):
        self.x = self.x + x_delta
        self.y = self.y + y_delta

    @property
    def height(self):
        return to_float(self.cell_value('Height'))

    @height.setter
    def height(self, value: float or str):
        self.set_cell_value('Height', str(value))

    @property
    def width(self):
        return to_float(self.cell_value('Width'))

    @width.setter
    def width(self, value: float or str):
        self.set_cell_value('Width', str(value))

    @property
    def center_x_y(self):
        x = self.x - (self.width/2)
        y = self.y + (self.height/2)
        return x, y

    @staticmethod
    def clear_all_text_from_xml(x: Element):
        x.text = ''
        x.tail = ''
        for i in x:
            Shape.clear_all_text_from_xml(i)

    @property
    def text(self):
        # return contents of Text element, or Master shape (if referenced), or empty string
        text_element = self.xml.find(f"{namespace}Text")

        if isinstance(text_element, Element):
            return "".join(text_element.itertext())  # get all text from <Text> sub elements
        elif self.master_page_ID:
            return self.master_shape.text  # get text from master shape
        return ""

    @text.setter
    def text(self, value):
        text_element = self.xml.find(f"{namespace}Text")
        if isinstance(text_element, Element):  # if there is a Text element then clear out and set contents
            Shape.clear_all_text_from_xml(text_element)
            text_element.text = value
        # todo: create new Text element if not found

    def sub_shapes(self) -> List[Shape]:
        """Get child/sub shapes contained by a Shape

        :returns: list of Shape objects
        :rtype: List[Shape]
        """
        shapes = list()
        # for each shapes tag, look for Shape objects
        # self can be either a Shapes or a Shape
        # a Shapes has a list of Shape
        # a Shape can have 0 or 1 Shapes (1 if type is Group)

        if self.shape_type == 'Group':
            parent_element = self.xml.find(f"{namespace}Shapes")
        else:  # a Shapes
            parent_element = self.xml
        if parent_element:
            shapes = [Shape(xml=shape, parent=self, page=self.page) for shape in parent_element]
        else:
            shapes = []
        return shapes

    def get_max_id(self):
        max_id = int(self.ID)
        if self.shape_type == 'Group':
            for shape in self.sub_shapes():
                new_max = shape.get_max_id()
                if new_max > max_id:
                    max_id = new_max
        return max_id

    def find_shape_by_id(self, shape_id: str) -> Shape:  # returns Shape
        """
        Recursively search for a shape, based on a known shape_id, and return a single Shape

        :param shape_id:
        :return: VisooFile.Shape
        """
        # recursively search for shapes by text and return first match
        for shape in self.sub_shapes():  # type: Shape
            if shape.ID == shape_id:
                return shape
            if shape.shape_type == 'Group':
                found = shape.find_shape_by_id(shape_id)
                if found:
                    return found

    def find_shapes_by_id(self, shape_id: str) -> List[Shape]:
        # recursively search for shapes by ID and return all matches
        found = list()
        for shape in self.sub_shapes():  # type: Shape
            if shape.ID == shape_id:
                found.append(shape)
            if shape.shape_type == 'Group':
                sub_found = shape.find_shapes_by_id(shape_id)
                if sub_found:
                    found.extend(sub_found)
        return found  # return list of matching shapes

    def find_shapes_by_master(self, master_page_ID: str, master_shape_ID: str) -> List[Shape]:
        # recursively search for shapes by master ID and return all matches
        found = list()
        for shape in self.sub_shapes():  # type: Shape
            if shape.master_shape_ID == master_shape_ID and shape.master_page_ID == master_page_ID:
                found.append(shape)
            if shape.shape_type == 'Group':
                sub_found = shape.find_shapes_by_master(master_shape_ID, master_shape_ID)
                if sub_found:
                    found.extend(sub_found)
        return found  # return list of matching shapes

    def find_shape_by_text(self, text: str) -> Shape:  # returns Shape
        # recursively search for shapes by text and return first match
        for shape in self.sub_shapes():  # type: Shape
            if text in shape.text:
                return shape
            if shape.shape_type == 'Group':
                found = shape.find_shape_by_text(text)
                if found:
                    return found

    def find_shapes_by_text(self, text: str, shapes: List[Shape] = None) -> List[Shape]:
        # recursively search for shapes by text and return all matches
        if not shapes:
            shapes = list()
        for shape in self.sub_shapes():  # type: Shape
            if text in shape.text:
                shapes.append(shape)
            if shape.shape_type == 'Group':
                found = shape.find_shapes_by_text(text)
                if found:
                    shapes.extend(found)
        return shapes

    def find_shape_by_property_label(self, property_label: str) -> Shape:  # returns Shape
        # recursively search for shapes by property name and return first match
        for shape in self.sub_shapes():  # type: Shape
            if property_label in shape.data_properties.keys():
                return shape
            if shape.shape_type == 'Group':
                found = shape.find_shape_by_property_label(property_label)
                if found:
                    return found

    def find_shapes_by_property_label(self, property_label: str, shapes: List[Shape] = None) -> List[Shape]:
        # recursively search for shapes by property name and return all matches
        if not shapes:
            shapes = list()
        for shape in self.sub_shapes():  # type: Shape
            if property_label in shape.data_properties.keys():
                shapes.append(shape)
            if shape.shape_type == 'Group':
                found = shape.find_shapes_by_property_label(property_label)
                if found:
                    shapes.extend(found)
        return shapes

    def apply_text_filter(self, context: dict):
        # check text against all context keys
        text = self.text
        for key in context.keys():
            r_key = "{{" + key + "}}"
            text = text.replace(r_key, str(context[key]))
        self.text = text

        for s in self.sub_shapes():
            s.apply_text_filter(context)

    def find_replace(self, old: str, new: str):
        # find and replace text in this shape and sub shapes
        text = self.text
        self.text = text.replace(old, new)

        for s in self.sub_shapes():
            s.find_replace(old, new)

    def remove(self):
        self.parent.xml.remove(self.xml)

    def append_shape(self, append_shape: Shape):
        # insert shape into shapes tag, and return updated shapes tag
        id_map = self.page.vis.increment_shape_ids(append_shape.xml, self.page)
        self.page.vis.update_ids(append_shape.xml, id_map)
        self.xml.append(append_shape.xml)

    @property
    def connects(self):
        # get list of connect items linking shapes
        connects = list()
        for c in self.page.connects:
            if self.ID in [c.shape_id, c.connector_shape_id]:
                connects.append(c)
        return connects

    @property
    def connected_shapes(self):
        # return a list of connected shapes
        shapes = list()
        for c in self.connects:
            if c.connector_shape_id != self.ID:
                shapes.append(self.page.find_shape_by_id(c.connector_shape_id))
            if c.shape_id != self.ID:
                shapes.append(self.page.find_shape_by_id(c.shape_id))
        return shapes
