import zipfile
import shutil

import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

import xml.dom.minidom as minidom   # minidom used for prettyprint


class VisioFile:
    def __init__(self, filename):
        self.filename = filename
        self.directory = f"./{filename.rsplit('.', 1)[0]}"
        self.pages = dict()   # populated by open_vsdx_file()
        self.page_max_ids = dict()  # maximum shape id, used to add new shapes with a unique Id

    @staticmethod
    def pretty_print_element(xml: Element) -> str:
        return minidom.parseString(ET.tostring(xml)).toprettyxml()

    def open_vsdx_file(self) -> dict:  # returns a dict of each page as ET with filename as key
        with zipfile.ZipFile(self.filename, "r") as zip_ref:
            zip_ref.extractall(self.directory)

        # load each page file into an ElementTree object
        self.pages = self.load_pages()

        return self.pages

    def load_pages(self):
        rel_dir = '{}/visio/pages/_rels/'.format(self.directory)
        page_dir = '{}/visio/pages/'.format(self.directory)

        rels = file_to_xml(rel_dir + 'pages.xml.rels').getroot()
        relid_page_dict = {}
        for rel in rels:
            rel_id=rel.attrib['Id']
            page_file = rel.attrib['Target']
            relid_page_dict[rel_id] = page_file

        pages = file_to_xml(page_dir + 'pages.xml').getroot()  # this contains a list of pages with rel_id and filename
        page_dict = {}  # dict with filename as index

        for page in pages:  # type: Element
            rel_id = page[1].attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            page_filename = relid_page_dict.get(rel_id, None)
            page_path = page_dir + page_filename
            page_dict[page_path] = file_to_xml(page_path)
            self.page_max_ids[page_path] = 0  # initialise page_max_ids dict for each page

        return page_dict

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
        # Remove extracted folder again
        shutil.rmtree(self.directory)

    def save_vsdx(self, new_filename=None):
        # write the pages to file
        for key in self.pages.keys():
            xml_to_file(self.pages[key], key)

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


def file_to_xml(filename: str) -> ET.ElementTree:
    tree = ET.parse(filename)
    return tree


def xml_to_file(xml: ET.ElementTree, filename: str):
    xml.write(filename)
