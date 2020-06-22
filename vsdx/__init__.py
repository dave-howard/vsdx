from __future__ import annotations
import zipfile
import shutil

import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

import xml.dom.minidom as minidom   # minidom used for prettyprint

namespace = "{http://schemas.microsoft.com/office/visio/2012/main}"  # visio file name space


# utility functions
def to_float(val: str):
    try:
        return float(val)
    except ValueError:
        return 0.0


class VisioFile:
    def __init__(self, filename):
        self.filename = filename
        self.directory = f"./{filename.rsplit('.', 1)[0]}"
        self.pages = dict()   # populated by open_vsdx_file()
        self.page_objects = list()  # list of Page objects
        self.page_max_ids = dict()  # maximum shape id, used to add new shapes with a unique Id
        self.open_vsdx_file()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close_vsdx()

    @staticmethod
    def pretty_print_element(xml: Element) -> str:
        return minidom.parseString(ET.tostring(xml)).toprettyxml()

    def open_vsdx_file(self) -> dict:  # returns a dict of each page as ET with filename as key
        with zipfile.ZipFile(self.filename, "r") as zip_ref:
            zip_ref.extractall(self.directory)

        # load each page file into an ElementTree object
        self.pages = self.load_pages()

        # todo: is this needed? remove
        #for filename, page in self.pages.items():
        #    p = VisioFile.Page(page, filename, 'no name', self)

        return self.pages

    def load_pages(self):
        rel_dir = '{}/visio/pages/_rels/'.format(self.directory)
        page_dir = '{}/visio/pages/'.format(self.directory)

        rels = file_to_xml(rel_dir + 'pages.xml.rels').getroot()
        #print(VisioFile.pretty_print_element(rels))  # rels contains map from filename to Id
        relid_page_dict = {}
        #relid_page_name = {}
        for rel in rels:
            rel_id=rel.attrib['Id']
            page_file = rel.attrib['Target']
            relid_page_dict[rel_id] = page_file
            #relid_page_name[rel_id] = page_name

        pages = file_to_xml(page_dir + 'pages.xml').getroot()  # this contains a list of pages with rel_id and filename
        #print(VisioFile.pretty_print_element(pages))  # pages contains Page name, width, height, mapped to Id
        page_dict = {}  # dict with filename as index

        for page in pages:  # type: Element
            rel_id = page[1].attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            page_name = page.attrib['Name']
            print(f"page_name:{page_name}")
            page_filename = relid_page_dict.get(rel_id, None)
            page_path = page_dir + page_filename
            page_dict[page_path] = file_to_xml(page_path)
            self.page_max_ids[page_path] = 0  # initialise page_max_ids dict for each page

            self.page_objects.append(VisioFile.Page(file_to_xml(page_path), page_path, page_name, self))

        return page_dict

    def get_page(self, n: int):
        try:
            # todo: also add get_page_by_name()
            return self.page_objects[n]
        except IndexError:
            return None

    def get_page_names(self):
        return [p.name for p in self.page_objects]

    def get_page_by_name(self, name: str):
        for p in self.page_objects:
            if p.name == name:
                return p

    def get_shapes(self, page_path) -> ET:
        page = self.pages[page_path]  # type: Element
        shapes = None
        # takes pages as an ET and returns a ET containing shapes
        for e in page.getroot():  # type: Element
            if 'Shapes' in e.tag:
                shapes = e
                for shape in e:  # type: Element
                    id = int(self.get_shape_id(shape))
                    max_id = self.page_max_ids[page_path]
                    if id > max_id:
                        self.page_max_ids[page_path] = id
        return shapes

    def get_sub_shapes(self, shape: Element, nth=1):
        for e in shape:
            if 'Shapes' in e.tag:
                nth -= 1
                if not nth:
                    return e

    @staticmethod
    def get_shape_location(shape: Element) -> (float, float):
        x, y = 0.0, 0.0
        for cell in shape:  # type: Element
            if 'Cell' in cell.tag:
                if cell.attrib.get('N'):
                    if cell.attrib['N'] == 'PinX':
                        x = float(cell.attrib['V'])
                    if cell.attrib['N'] == 'PinY':
                        y = float(cell.attrib['V'])
        return x, y

    @staticmethod
    def set_shape_location(shape: Element, x: float, y: float):
        for cell in shape:  # type: Element
            if 'Cell' in cell.tag:
                if cell.attrib.get('N'):
                    if cell.attrib['N'] == 'PinX':
                        cell.attrib['V'] = str(x)
                    if cell.attrib['N'] == 'PinY':
                        cell.attrib['V'] = str(y)

    @staticmethod
    def get_shape_text(shape: ET) -> str:
        text = None
        for t in shape:  # type: Element
            if 'Text' in t.tag:
                if t.text:
                    text = t.text
                else:
                    text = t[0].tail
        return text if text else ""

    @staticmethod
    def set_shape_text(shape: ET, text: str):
        for t in shape:  # type: Element
            if 'Text' in t.tag:
                if t.text:
                    t.text = text
                else:
                    t[0].tail = text

    # context = {'customer_name':'codypy.com', 'year':2020 }
    # example shape text "For {{customer_name}}  (c){{year}}" -> "For codypy.com (c)2020"
    @staticmethod
    def apply_text_context(shapes: Element, context: dict):
        for shape in shapes:  # type: Element
            if 'Shapes' in shape.tag:  # then this is a Shapes container
                VisioFile.apply_text_context(shape, context)  # recursive call
            if 'Shape' in shape.tag:
                # check text against all context keys
                for key in context.keys():
                    text = VisioFile.get_shape_text(shape)
                    r_key = "{{" + key + "}}"
                    if r_key in text:
                        new_text = text.replace(r_key, str(context[key]))
                        VisioFile.set_shape_text(shape, new_text)

    @staticmethod
    def get_shape_id(shape: ET) -> str:
        return shape.attrib['ID']

    def copy_shape(self, shape: Element, page: ET, page_path: str) -> ET:
        # insert shape into first Shapes tag in page
        new_shape = ET.fromstring(ET.tostring(shape))
        for shapes_tag in page.getroot():  # type: Element
            if 'Shapes' in shapes_tag.tag:
                id_map = self.increment_shape_ids(shape, page_path)
                self.update_ids(new_shape, id_map)
                shapes_tag.append(new_shape)
        return new_shape

    def insert_shape(self, shape: Element, shapes: Element, page: ET, page_path: str) -> ET:
        # insert shape into shapes tag, and return updated shapes tag
        id_map = self.increment_shape_ids(shape, page_path)
        self.update_ids(shape, id_map)
        shapes.append(shape)
        return shapes

    def increment_shape_ids(self, shape: Element, page_path: str, id_map: dict=None):
        if id_map is None:
            id_map = dict()
        self.set_new_id(shape, page_path, id_map)
        for e in shape:  # type: Element
            if 'Shapes' in e.tag:
                self.increment_shape_ids(e, page_path, id_map)
            if 'Shape' in e.tag:
                self.set_new_id(e, page_path, id_map)
        return id_map

    def set_new_id(self, element: Element, page_path: str, id_map: dict):
        if element.attrib.get('ID'):
            current_id = element.attrib['ID']
            max_id = self.page_max_ids[page_path] + 1
            id_map[current_id] = max_id  # record mappings
            element.attrib['ID'] = str(max_id)
            self.page_max_ids[page_path] = max_id
        else:
            print(f"no ID attr in {element.tag}")

    def update_ids(self, shape: Element, id_map: dict):
        # update: <ns0:Cell F="Sheet.15! replacing 15 with new id using prepopulated id_map
        # cycle through shapes looking for Cell tag inside a Shape tag, which may be inside a Shapes tag
        for e in shape:
            if 'Shapes' in e.tag:
                self.update_ids(e, id_map)
            if 'Shape' in e.tag:
                # look for Cell elements
                for c in e:  # type: Element
                    if 'Cell' in c.tag:
                        if c.attrib.get('F'):
                            f = str(c.attrib['F'])
                            if f.startswith("Sheet."):
                                # update sheet refs with new ids
                                id = f.split('!')[0].split('.')[1]
                                new_id = id_map[id]
                                new_f = f.replace(f'Sheet.{id}',f'Sheet.{new_id}')
                                c.attrib['F'] = new_f
        return shape

    def close_vsdx(self):
        try:
            # Remove extracted folder
            shutil.rmtree(self.directory)
        except FileNotFoundError:
            pass

    def save_vsdx(self, new_filename=None):
        # write the pages to file
        #for key in self.pages.keys():
        #    xml_to_file(self.pages[key], key)

        for page in self.page_objects:  # type: VisioFile.Page
            xml_to_file(page.xml, page.filename)

        # wrap up files into zip and rename to vsdx
        base_filename = self.filename[:-5]  # remove ".vsdx" from end
        shutil.make_archive(base_filename, 'zip', self.directory)
        if not new_filename:
            shutil.move(base_filename + '.zip', base_filename + '_new.vsdx')
        else:
            if new_filename[-5:] != '.vsdx':
                new_filename += '.vsdx'
            shutil.move(base_filename + '.zip', new_filename)
        self.close_vsdx()

    class Cell:
        def __init__(self, xml: Element, shape: VisioFile.Shape):
            self.xml = xml
            self.shape = shape

        @property
        def value(self):
            return self.xml.attrib.get('V')

        @property
        def name(self):
            return self.xml.attrib.get('N')

        @property
        def func(self):  # assume F stands for function, i.e. F="Width*0.5"
            return self.xml.attrib.get('F')

        def __repr__(self):
            return f"Cell: name={self.name} val={self.value} func={self.func}"

    class Shape:  # or page
        def __init__(self, xml: Element, parent_xml: Element, page: VisioFile.Page):
            self.xml = xml
            self.parent_xml = parent_xml
            self.tag = xml.tag
            self.ID = xml.attrib['ID'] if xml.attrib.get('ID') else None
            self.type = xml.attrib['Type'] if xml.attrib.get('Type') else None
            self.page = page

            # get Cells in Shape
            self.cells = dict()
            for e in self.xml:
                if e.tag == namespace+"Cell":
                    cell = VisioFile.Cell(xml=e, shape=self)
                    self.cells[cell.name] = cell

        def __repr__(self):
            return f"<Shape tag={self.tag} ID={self.ID} type={self.type} text='{self.text}' >"

        def cell_value(self, name: str):
            cell = self.cells.get(name)
            return cell.value if cell else None

        @property
        def x(self):
            return to_float(self.cell_value('PinX'))

        @property
        def y(self):
            return to_float(self.cell_value('PinY'))

        @property
        def height(self):
            return to_float(self.cell_value('Height'))

        @property
        def width(self):
            return to_float(self.cell_value('Width'))

        @property
        def text(self):
            text = None
            for t in self.xml:  # type: Element
                if 'Text' in t.tag:
                    if t.text:
                        text = t.text
                    else:
                        try:
                            text = t[0].tail if len(t) else ""
                        except IndexError:
                            print(f"Error getting t[0] where t={VisioFile.pretty_print_element(t)}")
            return text.replace('\n','') if text else ""

        @text.setter
        def text(self, value):
            for t in self.xml:  # type: Element
                if 'Text' in t.tag:
                    if t.text:
                        t.text = value
                    else:
                        t[0].tail = value

        def sub_shapes(self):
            shapes = list()
            # for each shapes tag, look for Shape objects
            for e in self.xml:  # type: Element
                #print(f"{e.tag}")
                if e.tag == namespace+'Shapes':
                    for shape in e:  # type: Element
                        if shape.tag == namespace+'Shape':
                            shapes.append(VisioFile.Shape(shape, e, self.page))
                if e.tag == namespace+'Shape':
                    shapes.append(VisioFile.Shape(e, self.xml, self.page))
            return shapes

        def find_shape_by_text(self, text: str) -> VisioFile.Shape:  # returns Shape
            for shape in self.sub_shapes():  # type: VisioFile.Shape
                if text in shape.text:
                    return shape
                if shape.type == 'Group':
                    found = shape.find_shape_by_text(text)
                    if found:
                        return found

        def apply_text_filter(self, context: dict):
            # check text against all context keys
            for key in context.keys():
                text = self.text
                r_key = "{{" + key + "}}"
                if r_key in text:
                    new_text = text.replace(r_key, str(context[key]))
                    self.text = new_text
            for s in self.sub_shapes():
                s.apply_text_filter(context)

        def remove(self):
            self.parent_xml.remove(self.xml)

        def append_shape(self, append_shape: VisioFile.Shape):
            # insert shape into shapes tag, and return updated shapes tag
            id_map = self.page.vis.increment_shape_ids(append_shape.xml, self.page.filename)
            self.page.vis.update_ids(append_shape.xml, id_map)
            self.xml.append(append_shape.xml)

    class Page:
        def __init__(self, xml: ET.ElementTree, filename: str, page_name: str, vis: VisioFile):
            self._xml = xml
            self.filename = filename
            self.name = page_name
            self.vis = vis

        def __repr__(self):
            return f"<Page name={self.name} file={self.filename} >"

        @property
        def xml(self):
            return self._xml

        @property
        def shapes(self):
            # list of Shape objects in Page
            page_shapes = list()
            for shape in self.xml.getroot():
                page_shapes.append(VisioFile.Shape(shape, self.xml, self))
            return page_shapes

        def apply_text_context(self, context: dict):
            for s in self.shapes:
                s.apply_text_filter(context)


def file_to_xml(filename: str) -> ET.ElementTree:
    tree = ET.parse(filename)
    return tree


def xml_to_file(xml: ET.ElementTree, filename: str):
    xml.write(filename)
