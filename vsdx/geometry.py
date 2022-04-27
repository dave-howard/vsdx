from __future__ import annotations
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

import vsdx

namespace = "{http://schemas.microsoft.com/office/visio/2012/main}"  # visio file name space


class Geometry:
    """ class to represent, and manipulate, the geometry of a shape"""
    def __init__(self, xml: Element, shape: vsdx.Shape):
        # get shape master geometry, and append/overwrite with actual shape instance data

        self.xml = xml  # expect an Element of Section with attr N='Geometry'
        self.cells = list()  #list of cells directly under Geometry section
        self.rows = dict()  # and list of rows with type(T) and index(IX), each containing a list of Cells
        self.shape = shape

        if shape.master_shape and shape.master_shape.geometry:
            self.cells = shape.master_shape.geometry.cells

        for cell in self.xml.findall(f"{namespace}Cell"):
            self.cells.append(GeometryCell(parent=self, xml=cell))

        if shape.master_shape and shape.master_shape.geometry:
            self.rows = shape.master_shape.geometry.rows  # type: dict
        for row in self.xml.findall(f"{namespace}Row"):
            index = row.attrib.get('IX')
            g_row = GeometryRow(geometry=self, xml=row, master_geometry_row=self.rows.get(index))
            self.rows[g_row.index] = g_row
            if g_row.del_bool:  # remove if master row over-ridden with a  deleted item
                del self.rows[g_row.index]

    def start_pos(self) -> tuple:
        # find start position of shape based on first MoveTo or RelMoveTo row in geometry
        for row in self.rows.values():  # type: GeometryRow
            if str(row.row_type).lower() == 'moveto':
                return row.x, row.y
            if str(row.row_type.lower()=='relmoveto'):
                # todo: find actual x,y based on shape width/height and relmoveto x,y
                return self.shape.x, self.shape.y

    def move(self, x_delta: float, y_delta: float):
        # update any absolute references to co-ordinates
        for r in self.rows.values():  # type: GeometryRow
            print(f"r={type(r)} {r}")
            if r.row_type.lower() in ['moveto', 'lineto']:  # todo: include other absolute row types
                r.x = r.x + x_delta if type(r.x) is float else None
                r.y = r.y + y_delta if type(r.y) is float else None
                print(f"r={type(r)} {r} after move {x_delta}, {y_delta}")

    def set_move_to(self, x: int, y: int, move_to_index: int=0):
        move_tos = [r for r in self.rows.values() if r.row_type.lower() == 'moveto']
        #print(f"move_tos={move_tos}")
        if len(move_tos) > move_to_index:
            move_to = move_tos[move_to_index]  # type: GeometryRow
            if move_to.geometry.shape.master_page_ID != self.shape.master_page_ID:
                move_to = GeometryRow(geometry=self, xml=None, master_geometry_row=move_to, T='MoveTo', IX=move_to.index)
                print(f"set_move_to() created: {move_to}")
            move_to.x = x
            move_to.y = y
            #print(f"move_to[{move_to_index}]={move_to.x},{move_to.y}")

    def set_line_to(self, x: int, y: int, line_to_index: int=0):
        line_tos = [r for r in self.rows.values() if r.row_type.lower() == 'lineto']
        #print(f"line_tos={line_tos}")
        if len(line_tos) > line_to_index:
            line_to = line_tos[line_to_index]  # type: GeometryRow
            if line_to.geometry.shape.master_page_ID != self.shape.master_page_ID:
                line_to = GeometryRow(geometry=self, xml=None, master_geometry_row=line_to, T='LineTo', IX=line_to.index)
                print(f"set_line_to() created: {line_to}")
            line_to.x = x
            line_to.y = y
            #print(f"line_to[{line_to_index}]={line_to.x},{line_to.y}")

    def __repr__(self):
        s = f"Geometry: {self.cells} {[(r.row_type, r.index, r.x,r.y) for r in self.rows.values()]}"
        s += f"\nGeometry: {vsdx.pretty_print_element(self.xml)}"
        return s


class GeometryRow:
    """A row with type(T) and index(IX), each containing a list of Cells"""
    """See: https://docs.microsoft.com/en-us/office/client-developer/visio/row-element-geometry-sectionvisio-xml """
    def __init__(self, geometry: Geometry, xml: Element, master_geometry_row: GeometryRow, T=None, IX=None):
        self.geometry = geometry  # parent of this row
        self.xml = xml if type(xml) is Element else self.create_row_xml(T, str(IX))
        # Create a dictionary of each Cell element, indexed by name
        self.cells = master_geometry_row.cells if master_geometry_row else dict()
        # add/overwrite cells values with master as basis id present
        for cell in self.xml.findall(f"{namespace}Cell"):
            g_cell = GeometryCell(parent=self, xml=cell)
            self.cells[g_cell.name] = g_cell

    def create_row_xml(self, T: str, IX: str):
        if T and IX:
            # Create new row xml
            row = ET.fromstring(f'<Row xmlns="{namespace[1:-1]}" T="{T}" IX="{IX}" />')
            # get all indexes
            indexes = [x.attrib.get('IX') for x in self.geometry.xml.findall(f"{namespace}Row")]
            if IX in indexes:
                row = None  # todo: replace existing row with new one
            else:
                indexes.append(IX)
                indexes.sort()
                self.geometry.xml.insert(indexes.index(IX), row)

            self.geometry.rows[IX] = self
            return row

    @property
    def row_type(self):
        return self.xml.attrib.get('T')

    @row_type.setter
    def row_type(self, value):
        self.xml.attrib['T'] = str(value)

    @property
    def index(self):
        return self.xml.attrib.get('IX')

    @index.setter
    def index(self, value):
        self.xml.attrib['IX'] = str(value)

    @property
    def x(self):
        x_cell = self.cells.get('X')
        return float(x_cell.value) if x_cell else None

    @x.setter
    def x(self, value):
        x_cell = self.cells.get('X')  # type: GeometryCell
        if not x_cell or (type(x_cell.parent) is GeometryRow and x_cell.parent.geometry.shape.master_page_ID!=self.geometry.shape.master_page_ID):
            # create new cell if none exists, or if existing cell is from master shape
            x_cell = GeometryCell(parent=self, xml=None, name='X', value=value)
            print(f"x_cell={x_cell}")
        x_cell.value = value

    @property
    def y(self):
        y_cell = self.cells.get('Y')
        return float(y_cell.value) if y_cell else None

    @y.setter
    def y(self, value):
        y_cell = self.cells.get('Y')
        if not y_cell or (type(y_cell.parent) is GeometryRow and y_cell.parent.geometry.shape.master_page_ID != self.geometry.shape.master_page_ID):
            # create new cell if none exists, or if existing cell is from master shape
            y_cell = GeometryCell(parent=self, xml=None, name='Y', value=value)
            print(f"y_cell={y_cell}")
        y_cell.value = value

    @property
    def del_bool(self):
        # Specifies whether a row that would otherwise be inherited from a master shape has been deleted.
        return self.xml.attrib.get('Del')

    @del_bool.setter
    def del_bool(self, value):
        if value:
            self.xml.attrib['Del'] = 1  # set to 1 if truthy
        else:
            del self.xml.attrib['Del']  # remove attribute if falsy

    def __repr__(self):
        s = f"Row[{self.index}] del:{self.del_bool}: {self.row_type}={self.cells}"
        return s


class GeometryCell:
    """class to represent a Cell element, a name value pair. This may be a child of Geometry or of GeometryRow"""
    def __init__(self, parent: GeometryRow or Geometry, xml: Element, name: str = None, value: str = None):
        self.parent = parent
        self.parent_xml = parent.xml
        self.xml = xml if type(xml) is Element else self.create_cell_xml(name)
        if name:
            self.name = name
        if value:
            self.value = value

    def create_cell_xml(self, name: str):
        # Create new row xml
        cell = ET.fromstring(f'<Cell xmlns="{namespace[1:-1]}"  />')
        self.parent_xml.append(cell)
        self.parent.cells[name] = self
        return cell

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

    @name.setter
    def name(self, value: str):
        self.xml.attrib['N'] = str(value)

    @property
    def func(self):  # assume F stands for function, i.e. F="Width*0.5"
        return self.xml.attrib.get('F')

    def __repr__(self):
        s = f"{self.name}={self.value}"
        if self.func:
            s += f" func={self.func}"
        return s
