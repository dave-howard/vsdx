from __future__ import annotations

import zipfile
import shutil
import os
import re
import io

from jinja2 import Template

from typing import List
from typing import Optional

import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element
import xml.dom.minidom as minidom   # minidom used for prettyprint

import vsdx
from .pages import Page
from .pages import PagePosition

from vsdx import Shape

from vsdx import namespace
from vsdx import ext_prop_namespace
from vsdx import vt_namespace
from vsdx import r_namespace
from vsdx import document_rels_namespace
from vsdx import cont_types_namespace

ET.register_namespace('', namespace[1:-1])
ET.register_namespace('', ext_prop_namespace[1:-1])
ET.register_namespace('vt', vt_namespace[1:-1])
ET.register_namespace('r', r_namespace[1:-1])
ET.register_namespace('', document_rels_namespace[1:-1])
ET.register_namespace('', cont_types_namespace[1:-1])


def file_to_xml(filename: str, zip_file_contents: dict = None) -> ET.ElementTree:
    """Import a file as an ElementTree"""
    try:
        if zip_file_contents and zip_file_contents.get(filename):
            print(f'loading {filename} from zip_file_contents')
            content : io.BytesIO = zip_file_contents[filename]
            print(f'content type={type(content)}')
            tree = ET.parse(io.BytesIO(content.getvalue()))
            return tree
        tree = ET.parse(filename)
        return tree
    except FileNotFoundError:
        pass  # return None


def xml_to_file(xml: ET.ElementTree, filename: str, zip_file_contents: dict = None):
    """Save an ElementTree to a file or zip_file_contents"""
    if zip_file_contents:
        file : io.BytesIO = io.BytesIO()
        print(f'befoe writing {filename} file={file} {len(file.getvalue())}')
        xml.write(file, xml_declaration=True, method='xml', encoding='UTF-8')
        print(f'after writing {filename} file={file} {len(file.getvalue())}')
        zip_file_contents[filename] = io.BytesIO(file.getvalue())
        print(f'saved {filename} to zip_file_contents\n{zip_file_contents[filename].getvalue()[:100]}')
    else:
        xml.write(filename, xml_declaration=True, method='xml', encoding='UTF-8')
        print(f'saved {filename} to file')


class VisioFileNotOpen(BaseException):
    """Error class to report when a VisioFile is attempted to be saved when no longer open"""
    pass


class VisioFile:
    """Represents a vsdx file

    :param filename: filename the :class:`VisioFile` was created from
    :type filename: str
    :param pages: a list of pages in the VisioFile
    :type pages: list of :class:`Page`
    :param master_pages: a list of master pages in the VisioFile
    :type master_pages: list of :class:`Page`
    """
    def __init__(self, filename, debug: bool = False):
        """VisioFile constructor

        :param filename: the vsdx file to load and create the VisioFile object from
        :type filename: str
        :param debug: enable/disable debugging
        :type debug: bool, default to False
        """
        self.debug = debug
        self.filename = filename
        if debug:
            print(f"VisioFile(filename={filename})")
        file_type = self.filename.split('.')[-1]  # last text after dot
        if not file_type.lower() == 'vsdx' and not file_type.lower() == 'vsdm':
            raise TypeError(f'Invalid File Type:{file_type}')

        self.directory = os.path.abspath(filename)[:-5]
        self.pages_xml = None  # type: ET.ElementTree
        self.pages_xml_rels = None  # type: ET.ElementTree
        self.content_types_xml = None  # type: ET.ElementTree
        self.app_xml = None  # type: ET.ElementTree
        self.document_xml = None  # type: ET.ElementTree
        self.document_xml_rels = None  # type: ET.ElementTree
        self.pages = list()  # type: List[Page]  # list of Page objects, populated by open_vsdx_file()
        self.masters_xml = None  # type: ET.ElementTree
        self.master_index = {}  # dict of master page info by item name e.g. 'Dynamic Connector'
        self.master_pages = list()  # type: List[Page]  # list of Page objects, populated by open_vsdx_file()
        self.file_open = False
        self.zip_file_contents = {}  # dict of file contents by file_path
        self.open_vsdx_file()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close_vsdx()

    @staticmethod
    def pretty_print_element(xml: Element) -> str:
        if type(xml) is Element:
            return minidom.parseString(ET.tostring(xml)).toprettyxml()
        elif type(xml) is ET.ElementTree:
            return minidom.parseString(ET.tostring(xml.getroot())).toprettyxml()
        else:
            return f"Not an Element. type={type(xml)}"

    def _load_zip_file_contents_to_memory(self):
        """Open zip file and create a dictionary of file like objects by file_path"""
        with zipfile.ZipFile(self.filename, "r") as zip_ref:
            for file_path in zip_ref.namelist():
                path = f"{self.directory}/{file_path}"
                print(f"file_path:{path}")
                if not path.endswith('/'):  # ignore directories
                    content = zip_ref.read(file_path)
                    self.zip_file_contents[path] = io.BytesIO(content)
                #print(f"path={path} content type={type(self.zip_file_contents.get(path))}")
        print(f'_load_zip_file_contents_to_memory() zip_file_contents={self.zip_file_contents.keys()}')

    def _save_zip_file_contents_to_disk(self, save_filename: str):
        """Save the zip_file_contents to disk"""
        print(f'_save_zip_file_contents_to_memory() zip_file_contents={self.zip_file_contents.keys()}')
        with zipfile.ZipFile(save_filename, "w") as zipf:
            for file_path, file_content in self.zip_file_contents.items(): # type: tuple(str, io.BytesIO)
                file_path_in_zip : str = file_path.replace(self.directory+'/', '')
                print(f'writing {file_path_in_zip} to zip file')
                try:
                    if file_path_in_zip.endswith('.xml') or file_path_in_zip.endswith('.rels'):
                        zipf.writestr(file_path_in_zip, file_content.read().decode('utf-8'))
                    else:
                        zipf.writestr(file_path_in_zip, file_content.read())
                except Exception as e:
                    print(f'Error writing {file_path_in_zip} to zip file: {e}')
                    raise e

    def open_vsdx_file(self):
        self._load_zip_file_contents_to_memory()
        with zipfile.ZipFile(self.filename, "r") as zip_ref:
            zip_ref.extractall(self.directory)

        # load each page file into an ElementTree object
        self.load_pages()
        self.load_master_pages()
        self.file_open = True

    def _pages_filename(self):
        page_dir = f'{self.directory}/visio/pages/'
        pages_filename = page_dir + 'pages.xml'  # pages.xml contains Page name, width, height, mapped to Id
        return pages_filename

    @property
    def _masters_folder(self):
        path = f"{self.directory}/visio/masters"
        return path

    def load_pages(self):
        rel_dir = f'{self.directory}/visio/pages/_rels/'
        page_dir = f'{self.directory}/visio/pages/'

        rel_filename = rel_dir + 'pages.xml.rels'
        rels = file_to_xml(rel_filename, self.zip_file_contents).getroot()  # rels contains page filenames
        self.pages_xml_rels = file_to_xml(rel_filename, self.zip_file_contents)  # store pages.xml.rels so pages can be added or removed
        if self.debug:
            print(f"Relationships({rel_filename})", VisioFile.pretty_print_element(rels))
        relid_page_dict = {}

        for rel in rels:
            rel_id = rel.attrib['Id']
            page_file = rel.attrib['Target']
            relid_page_dict[rel_id] = page_file

        pages_filename = self._pages_filename()  # pages contains Page name, width, height, mapped to Id
        pages = file_to_xml(pages_filename, self.zip_file_contents).getroot()  # this contains a list of pages with rel_id and filename
        self.pages_xml = file_to_xml(pages_filename, self.zip_file_contents)  # store xml so pages can be removed
        if self.debug:
            print(f"Pages({pages_filename})", VisioFile.pretty_print_element(pages))

        for page in pages:  # type: Element
            rel_id = page.find(f"{namespace}Rel").attrib[f"{r_namespace}id"]
            page_name = page.attrib['Name']

            page_path = page_dir + relid_page_dict.get(rel_id, None)
            page_id = page.attrib.get('ID')

            new_page = Page(file_to_xml(page_path, self.zip_file_contents), page_path, page_name, page_id, rel_id, self)
            # look for visio/pages/_rels/page3.xml.rels
            base_page_file_name = page_path.split('/')[-1]
            page_rels_path = rel_dir+base_page_file_name+'.rels'

            if os.path.exists(page_rels_path):
                new_page.rels_xml_filename = page_rels_path
                new_page.rels_xml = file_to_xml(page_rels_path, self.zip_file_contents)
            self.pages.append(new_page)

            if self.debug:
                print(f"Page({new_page.filename})", VisioFile.pretty_print_element(new_page.xml.getroot()))

        self.content_types_xml = file_to_xml(f'{self.directory}/[Content_Types].xml', self.zip_file_contents)
        # TODO: add correctness cross-check. Or maybe the other way round, start from [Content_Types].xml
        #       to get page_dir and other paths...

        self.app_xml = file_to_xml(f'{self.directory}/docProps/app.xml', self.zip_file_contents)  # note: files in docProps may be missing
        self.document_xml = file_to_xml(f'{self.directory}/visio/document.xml', self.zip_file_contents)
        self.document_xml_rels = file_to_xml(f'{self.directory}/visio/_rels/document.xml.rels', self.zip_file_contents)

    def load_master_pages(self):
        # get data from /visio/masters folder
        master_rel_path = f'{self.directory}/visio/masters/_rels/masters.xml.rels'

        master_rels_data = file_to_xml(master_rel_path, self.zip_file_contents)
        master_rels = master_rels_data.getroot() if master_rels_data else []
        if self.debug:
            print(f"Master Relationships({master_rel_path})", VisioFile.pretty_print_element(master_rels))

        # populate relid to master path
        relid_to_path = {}
        for rel in master_rels:
            master_id = rel.attrib.get('Id')
            master_path = f"{self.directory}/visio/masters/{rel.attrib.get('Target')}"  # get path from rel
            relid_to_path[master_id] = master_path

        # load masters.xml file
        masters_path = f'{self.directory}/visio/masters/masters.xml'
        masters_xml = file_to_xml(masters_path, self.zip_file_contents)  # contains more info about master page (i.e. Name, Icon)
        self.masters_xml = masters_xml.getroot() if masters_xml else []

        # for each master page, create the Page object
        for master in self.masters_xml:
            master_name = master.attrib.get('NameU') or master.attrib.get('Name') or 'Unknown'
            rel_id = master.find(f"{namespace}Rel").attrib[f"{r_namespace}id"]
            master_id = master.attrib['ID']
            master_unique_id = master.attrib.get('UniqueID')
            master_base_id = master.attrib.get('BaseID')

            master_path = relid_to_path[rel_id]

            master_page = Page(file_to_xml(master_path, self.zip_file_contents), master_path, master_name, master_id, rel_id, self)
            master_page.master_unique_id = master_unique_id
            master_page.master_base_id = master_base_id
            self.master_pages.append(master_page)
            self.master_index[master_name] = master_page  # index by master_name

            if self.debug:
                print(f"Master({master_path}, id={master_id})", VisioFile.pretty_print_element(master_page.xml.getroot()))

        return

    def get_page(self, n: int) -> Page:
        try:
            return self.pages[n]
        except IndexError:
            return None

    def get_page_names(self):
        return [p.name for p in self.pages]

    def get_page_by_name(self, name: str):
        """Get page from VisioFile with matching name

                :param name: The name of the required page
                :type name: str

                :return: :class:`Page` object representing the page (or None if not found)
                """
        for p in self.pages:
            if p.name == name:
                return p

    def get_master_page_by_id(self, id: str):
        """Get master page from VisioFile with matching ID.

        Referred by :attr:`Shape.master_ID`.

                :param id: The ID of the required master
                :type id: str

                :return: :class:`Page` object representing the master page (or None if not found)
                """
        for m in self.master_pages:
            if m.page_id == id:
                return m

    def remove_page_by_index(self, index: int):
        """Remove zero-based nth page from VisioFile object

        :param index: Zero-based index of the page
        :type index: int

        :return: None
        """

        # remove Page element from pages.xml file - zero based index
        if type(index) is int:
            page = self.pages_xml.find(f"{namespace}Page[{index+1}]")
            if isinstance(page, Element):
                self.pages_xml.getroot().remove(page)
                page = self.pages[index]  # type: Page

                # remove internal references to page
                self._remove_page_from_app_xml(page.name)

                # remove page<index>.xml file
                print(self.pages[index].filename)
                os.remove(self.pages[index].filename)
                del self.pages[index]


    def remove_page_by_name(self, page_name):
        """Remove first page from VisioFile object that matches the page_name

        :param page_name: page of page to delete
        :type page_name: str

        :return: None
        """

        # get index and then pass to remove_page_by_index() to perform deletion
        for p in self.pages:
            if p.name == page_name:
                self.remove_page_by_index(p.index_num)
                break  # exit after first match - delete only one page


    def _update_pages_xml_rels(self, new_page_filename: str) -> str:
        '''Updates the pages.xml.rels file with a reference to the new page and returns the new relid
        '''

        max_relid = max(self.pages_xml_rels.getroot(), key=lambda rel: int(rel.attrib['Id'][3:]), default=None)  # 'rIdXX' -> XX
        max_relid = int(max_relid.attrib['Id'][3:]) if max_relid is not None else 0
        new_page_relid = f'rId{max_relid + 1}'  # Most likely will be equal to len(self.pages)+1

        new_page_rel = {
            'Target': new_page_filename,
            'Type'  : 'http://schemas.microsoft.com/visio/2010/relationships/page',
            'Id'    : new_page_relid
        }
        self.pages_xml_rels.getroot().append(Element('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship', new_page_rel))

        return new_page_relid

    def _get_new_page_name(self, new_page_name: str) -> str:
        i = 1
        while new_page_name in self.get_page_names():
            new_page_name = f'{new_page_name}-{i}'  # Page-X-i
            i += 1

        return new_page_name

    def _get_max_page_id(self) -> int:
        page_with_max_id = max(self.pages_xml.getroot(), key=lambda page: int(page.attrib['ID']))
        max_page_id = int(page_with_max_id.attrib['ID'])

        return max_page_id

    def _get_index(self, *, index: int, page: Page or None):
        if type(index) is PagePosition:  # only update index if it is relative to source page
            if index == PagePosition.LAST:
                index = len(self.pages)
            elif index == PagePosition.FIRST:
                index = 0
            elif page:  # need page for BEFORE or AFTER
                orig_page_idx = self.pages.index(page)
                if index == PagePosition.BEFORE:
                    # insert new page at the original page's index
                    index = orig_page_idx
                elif index == PagePosition.AFTER:
                    # insert new page after the original page
                    index = orig_page_idx + 1
            else:
                index = len(self.pages)  # default to LAST if invalid Position/page combination

        return index

    def _add_content_types_override(self, part_name_path: str, content_type: str):
        content_types = self.content_types_xml.getroot()

        content_types_attribs = {
            'PartName': part_name_path,
            'ContentType': content_type,
        }
        override_element = Element(f'{cont_types_namespace}Override', content_types_attribs)
        # find existing elements with same content_type
        matching_overrides = content_types.findall(
            f'{cont_types_namespace}Override[@ContentType="{content_type}"]'
        )
        if len(matching_overrides):  # insert after similar elements
            idx = list(content_types).index(matching_overrides[-1])
            content_types.insert(idx+1, override_element)
        else:  # add at end of list
            content_types.append(override_element)

    def _update_content_types_xml(self, new_page_filename: str):
        # todo: use generic function above
        content_types = self.content_types_xml.getroot()

        content_types_attribs = {
            'PartName'   : f'/visio/pages/{new_page_filename}',
            'ContentType': 'application/vnd.ms-visio.page+xml'
        }
        content_types_element = Element(f'{cont_types_namespace}Override', content_types_attribs)

        # add the new element after the last such element
        # first find the index:
        all_page_overrides = content_types.findall(
            f'{cont_types_namespace}Override[@ContentType="application/vnd.ms-visio.page+xml"]'
        )
        idx = list(content_types).index(all_page_overrides[-1])

        # then add it:
        content_types.insert(idx+1, content_types_element)

    def document_rels(self) -> List[Element]:
        rels = self.document_xml_rels.findall(f'{document_rels_namespace}Relationship')
        return rels

    def _add_document_rel(self, rel_type: str, target: str):
        rel_ids = [int(str(r.attrib.get('Id')).replace('rId', '')) for r in self.document_rels()]
        new_rel = Element(f"{document_rels_namespace}Relationship",
                          {
                              "Id": f"rId{max(rel_ids)+1}",
                              "Type": rel_type,
                              "Target": target,
                          })
        self.document_xml_rels.getroot().append(new_rel)

    def _style_sheets(self) -> Element:
        # return StyleSheets element from document.xml
        return self.document_xml.getroot().find(f'{namespace}StyleSheets')

    def _get_styles_name_list(self) -> List[str]:
        return [s.attrib.get('Name','') for s in self._style_sheets().findall(f'{namespace}StyleSheet')]

    def _get_style_by_name(self, name: str) -> Element:
        stylesheet = self._style_sheets().find(f"{namespace}StyleSheet[@Name = '{name}']")
        return stylesheet

    def _get_style_by_id(self, ID: str) -> Element:
        stylesheet = self._style_sheets().find(f"{namespace}StyleSheet[@ID = '{ID}']")
        return stylesheet

    def _heading_pairs(self) -> Element:
        # return HeadingPairs element from app.xml
        return self.app_xml.getroot().find(f'{ext_prop_namespace}HeadingPairs')

    def _titles_of_parts(self) -> Element:
        # return TitlesOfParts element from app.xml
        return self.app_xml.getroot().find(f'{ext_prop_namespace}TitlesOfParts')

    def _titles_of_parts_list(self) -> List[str]:
        # return list of strings
        return [t.text for t in self._titles_of_parts().find(f".//{vt_namespace}vector")]

    def _add_titles_of_parts_item(self, title: str):
        titles = self._titles_of_parts()
        vector = titles.find(f".//{vt_namespace}vector")  # new variant appended to vector Element
        new_title = Element(f'{vt_namespace}lpstr', {})
        new_title.text = title
        vector.append(new_title)
        # add one to vector size, as we have added two new variant elements
        vector.attrib['size'] = str(int(vector.attrib.get('size', 0)) + 1)

    def _get_app_xml_value(self, name: str) -> str:
        variants = self._heading_pairs().findall(f".//{vt_namespace}variant")
        # find Pages in headings
        for index in range(len(variants)):
            v = variants[index]
            lpstr = v.find(f".//{vt_namespace}lpstr")
            if type(lpstr) is Element and lpstr.text == name:
                next_v = variants[index+1] if index < (len(variants)-1) else None  # next variant if there is one
                i4 = next_v.find(f".//{vt_namespace}i4") if type(next_v) is Element else None
                if type(i4) is Element:
                    return i4.text

    def _set_app_xml_value(self, name: str, value: str):
        variants = self._heading_pairs().findall(f".//{vt_namespace}variant")
        # find Pages in headings
        for index in range(len(variants)):
            v = variants[index]
            lpstr = v.find(f".//{vt_namespace}lpstr")
            if type(lpstr) is Element and lpstr.text == name:
                next_v = variants[index+1] if index < (len(variants)-1) else None  # next variant if there is one
                i4 = next_v.find(f".//{vt_namespace}i4") if type(next_v) is Element else None
                if type(i4) is Element:
                    i4.text = value
                    return
        # no matching variant found - so create new item and populate it
        vector = self._heading_pairs().find(f".//{vt_namespace}vector")  # new variant appended to vector Element
        name_variant = Element(f'{vt_namespace}variant', {})
        lpstr = Element(f'{vt_namespace}lpstr', {})
        lpstr.text = name
        name_variant.append(lpstr)
        vector.append(name_variant)
        i4_variant = Element(f'{vt_namespace}variant', {})
        i4 = Element(f'{vt_namespace}i4', {})
        i4.text = value
        i4_variant.append(i4)
        vector.append(i4_variant)
        # add two to vector size, as we have added two new variant elements
        vector.attrib['size'] = str(int(vector.attrib.get('size', 0)) + 2)

    def _add_page_to_app_xml(self, new_page_name: str):
        # todo: use _add_titles_of_parts_item()
        HeadingPairs = self._heading_pairs()
        i4 = HeadingPairs.find(f'.//{vt_namespace}i4')
        num_pages = int(i4.text)
        i4.text = str(num_pages+1)  # increment as page added

        TitlesOfParts = self.app_xml.getroot().find(f'{ext_prop_namespace}TitlesOfParts')
        vector = TitlesOfParts.find(f'{vt_namespace}vector')

        lpstr = Element(f'{vt_namespace}lpstr')
        lpstr.text = new_page_name
        vector.append(lpstr)  # add new lpstr element with new page name
        vector_size = int(vector.attrib['size'])
        vector.set('size', str(vector_size+1))  # increment as page added

    def _remove_page_from_app_xml(self, page_name: str):
        if self.app_xml is not None:
            print(f"_remove_page_from_app_xml()")
            HeadingPairs = self.app_xml.getroot().find(f'{ext_prop_namespace}HeadingPairs')
            i4 = HeadingPairs.find(f'.//{vt_namespace}i4')
            num_pages = int(i4.text)
            i4.text = str(num_pages-1)  # decrement as page removed

            TitlesOfParts = self.app_xml.getroot().find(f'{ext_prop_namespace}TitlesOfParts')
            vector = TitlesOfParts.find(f'{vt_namespace}vector')

            for lpstr in vector.findall(f'{vt_namespace}lpstr'):
                if lpstr.text == page_name:
                    vector.remove(lpstr)  # remove page from list of names
                    break

            vector_size = int(vector.attrib['size'])
            vector.set('size', str(vector_size-1))  # decrement as page removed

    def _create_page(
        self,
        *,
        new_page_xml_str: str,
        page_name: str,
        new_page_element: Element,
        index: int or PagePosition,
        source_page: Optional[Page] = None,
    ) -> Page:
        # Create visio\pages\pageX.xml file
        # Add to visio\pages\_rels\pages.xml.rels
        # Add to visio\pages\pages.xml
        # Add to [Content_Types].xml
        # Add to docProps\app.xml

        page_dir = f'{self.directory}/visio/pages/'  # TODO: better concatenation

        # create pageX.xml
        new_page_xml = ET.ElementTree(ET.fromstring(new_page_xml_str))
        new_page_filename = f'page{len(self.pages) + 1}.xml'
        new_page_path = page_dir+new_page_filename  # TODO: better concatenation

        # update pages.xml.rels - add rel for the new page
        # done by the caller

        # update pages.xml - insert the PageElement Element in it's correct location
        index = self._get_index(index=index, page=source_page)
        self.pages_xml.getroot().insert(index, new_page_element)

        # update [Content_Types].xml - insert reference to the new page
        self._update_content_types_xml(new_page_filename)

        # update app.xml, if it exists
        if self.app_xml:
            self._add_page_to_app_xml(page_name)

        # Update VisioFile object
        new_page = Page(new_page_xml, new_page_path, page_name, '', '', self)

        self.pages.insert(index, new_page)  # insert new page at defined index

        return new_page

    def add_page_at(self, index: int, name: Optional[str] = None) -> Page:
        """Add a new page at the specified index of the VisioFile

        :param index: zero-based index where the new page will be placed
        :type index: int

        :param name: The name of the new page
        :type name: str, optional

        :return: :class:`Page` object representing the new page
        """

        # Determine the new page's name
        new_page_name = self._get_new_page_name(name or f'Page-{len(self.pages) + 1}')

        # Determine the new page's filename
        new_page_filename = f'page{len(self.pages) + 1}.xml'

        # Add reference to the new page in pages.xml.rels and get new relid
        new_page_relid = self._update_pages_xml_rels(new_page_filename)

        # Create default empty page xml
        # TODO: figure out the best way to define this default pagesheet XML
        # For example, python-docx has a 'template.docx' file which is copied.
        new_pagesheet_attribs = {
            'FillStyle': '0',
            'LineStyle': '0',
            'TextStyle': '0'
        }
        new_pagesheet_element = Element(f'{namespace}PageSheet', new_pagesheet_attribs)
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'PageWidth', 'V':'8.26771653543307'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'PageHeight', 'V':'11.69291338582677'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'ShdwOffsetX', 'V':'0.1181102362204724'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'ShdwOffsetY', 'V':'-0.1181102362204724'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'PageScale', 'U':'MM', 'V':'0.03937007874015748'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'DrawingScale', 'U':'MM', 'V':'0.03937007874015748'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'DrawingSizeType', 'V':'0'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'DrawingScaleType', 'V':'0'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'InhibitSnap', 'V':'0'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'PageLockReplace', 'U':'BOOL', 'V':'0'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'PageLockDuplicate', 'U':'BOOL', 'V':'0'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'UIVisibility', 'V':'0'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'ShdwType', 'V':'0'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'ShdwObliqueAngle', 'V':'0'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'ShdwScaleFactor', 'V':'1'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'DrawingResizeType', 'V':'1'}))
        new_pagesheet_element.append(Element(f'{namespace}Cell', {'N':'PageShapeSplit', 'V':'1'}))

        new_page_attribs = {
            'ID'   : str(self._get_max_page_id() + 1),
            'NameU': new_page_name,
            'Name' : new_page_name,
        }
        new_page_element = Element(f'{namespace}Page', new_page_attribs)
        new_page_element.append(new_pagesheet_element)

        new_page_rel = {
            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id' : new_page_relid
        }
        new_page_element.append(Element(f'{namespace}Rel', new_page_rel))

        # create the new page
        new_page = self._create_page(
            new_page_xml_str = f"<?xml version='1.0' encoding='utf-8' ?><PageContents xmlns='{namespace[1:-1]}' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xml:space='preserve'/>",
            page_name = new_page_name,
            new_page_element = new_page_element,
            index = index
        )

        return new_page

    def add_page(self, name: Optional[str] = None) -> Page:
        """Add a new page at the end of the VisioFile

        :param name: The name of the new page
        :type name: str, optional

        :return: Page object representing the new page
        """

        return self.add_page_at(PagePosition.LAST, name)

    def copy_page(self, page: Page, *, index: Optional[int] = PagePosition.AFTER, name: Optional[str] = None) -> Page:
        """Copy an existing page and insert in VisioFile

        :param page: the page to copy
        :type page: Page
        :param index: the specific int or relation PagePosition location for new page
        :type index: int or PagePosition
        :param name: name of new page (note this may be altered if name already exists)
        :type name: str

        :return: the newly created page
        """
        # Determine the new page's name
        new_page_name = self._get_new_page_name(name or page.name)

        # Determine the new page's filename
        new_page_filename = f'page{len(self.pages) + 1}.xml'

        # Add reference to the new page in pages.xml.rels and get new relid
        new_page_relid = self._update_pages_xml_rels(new_page_filename)

        # Copy the source page and update relevant attributes
        page_element = self.pages_xml.find(f"{namespace}Page[@Name='{page.name}']")
        new_page_element = ET.fromstring(ET.tostring(page_element))

        new_page_element.attrib['ID']    = str(self._get_max_page_id() + 1)
        new_page_element.attrib['NameU'] = new_page_name
        new_page_element.attrib['Name']  = new_page_name
        new_page_element.find(f'{namespace}Rel').attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'] = new_page_relid

        # create the new page
        new_page = self._create_page(
            new_page_xml_str = ET.tostring(page.xml.getroot()),
            page_name = new_page_name,
            new_page_element = new_page_element,
            index = index,
            source_page = page,
        )

        # copy pageX.xml.rels if it exists
        # from testing, this does not actually seem to make a difference
        _, original_filename = os.path.split(page.filename)
        page_xml_rels_file = f'{self.directory}/visio/pages/_rels/{original_filename}.rels'  # TODO: better concatenation
        new_page_xml_rels_file = f'{self.directory}/visio/pages/_rels/{new_page_filename}.rels'  # TODO: better concatenation
        try:
            shutil.copy(page_xml_rels_file, new_page_xml_rels_file)
        except FileNotFoundError:
            pass

        return new_page

    # TODO: dead code - never used
    def get_sub_shapes(self, shape: Element, nth=1):
        for e in shape:
            if 'Shapes' in e.tag:
                nth -= 1
                if not nth:
                    return e

    @staticmethod
    def get_shape_location(shape: Element) -> (float, float):
        x, y = 0.0, 0.0
        cell_PinX = shape.find(f'{namespace}Cell[@N="PinX"]')  # type: Element
        cell_PinY = shape.find(f'{namespace}Cell[@N="PinY"]')
        x = float(cell_PinX.attrib['V'])
        y = float(cell_PinY.attrib['V'])

        return x, y

    @staticmethod
    def set_shape_location(shape: Element, x: float, y: float):
        cell_PinX = shape.find(f'{namespace}Cell[@N="PinX"]')  # type: Element
        cell_PinY = shape.find(f'{namespace}Cell[@N="PinY"]')
        cell_PinX.attrib['V'] = str(x)
        cell_PinY.attrib['V'] = str(y)

    @staticmethod
    # TODO: is this never used?
    def get_shape_text(shape: ET) -> str:
        # technically the below is not an exact replacement of the above...
        text = ""
        text_elem = shape.find(f"{namespace}Text")
        if text_elem is not None:
            text = "".join(text_elem.itertext())
        return text

    @staticmethod
    # TODO: is this never used?
    def set_shape_text(shape: ET, text: str):
        t = shape.find(f"{namespace}Text")  # type: Element
        if t is not None:
            if t.text:
                t.text = text
            else:
                t[0].tail = text

    # context = {'customer_name':'codypy.com', 'year':2020 }
    # example shape text "For {{customer_name}}  (c){{year}}" -> "For codypy.com (c)2020"
    @staticmethod
    def apply_text_context(shapes: Element, context: dict):

        def _replace_shape_text(shape: Element, context: dict):
            text = VisioFile.get_shape_text(shape)

            for key in context.keys():
                r_key = "{{" + key + "}}"
                text = text.replace(r_key, str(context[key]))
            VisioFile.set_shape_text(shape, text)

        for shape in shapes.findall(f"{namespace}Shapes"):
                VisioFile.apply_text_context(shape, context)  # recursive call
                _replace_shape_text(shape, context)

        for shape in shapes.findall(f"{namespace}Shape"):
            _replace_shape_text(shape, context)

    def jinja_render_vsdx(self, context: dict):
        """Transform a template VisioFile object using the Jinja language
        The method updates the VisioFile object loaded from the template file, so does not return any value
        Note: vsdx specific extensions are available such as `{% for item in list %}` statements with no `{% endfor %}`

        :param context: A dictionary containing values that can be accessed by the Jinja processor
        :type context: dict

        :return: None
        """
        # parse each shape in each page as Jinja2 template with context
        pages_to_remove = []  # list of pages to be removed after loop
        for page in self.pages:  # type: Page
            # check if page should be removed
            if VisioFile.jinja_page_showif(page, context):
                loop_shape_ids = list()
                for shapes_by_id in page._shapes:  # type: Shape
                    VisioFile.jinja_render_shape(shape=shapes_by_id, context=context, loop_shape_ids=loop_shape_ids)

                source = ET.tostring(page.xml.getroot(), encoding='unicode')
                source = VisioFile.unescape_jinja_statements(source)  # unescape chars like < and > inside {%...%}
                template = Template(source)
                output = template.render(context)
                page.xml = ET.ElementTree(ET.fromstring(output))  # create ElementTree from Element created from output

                # update loop shape IDs which have been duplicated by Jinja template
                page.set_max_ids()
                for shape_id in loop_shape_ids:
                    shapes_by_id = page._find_shapes_by_id(shape_id)  # type: List[Shape]
                    if shapes_by_id and len(shapes_by_id) > 1:
                        delta = 0
                        for shape in shapes_by_id[1:]:  # from the 2nd onwards - leaving original unchanged
                            # increment each new shape duplicated by the jinja loop
                            self.increment_sub_shape_ids(shape, page)
                            delta += shape.height  # automatically move each duplicate down
                            shape.move(0, -delta)  # move duplicated shapes so they are visible
            else:
                # note page to remove after this loop has completed
                pages_to_remove.append(page)
        # remove pages after processing
        for p in pages_to_remove:
            print(f"Removing page:'{p.name}' index:{p.index_num}")
            self.remove_page_by_index(p.index_num)

    @staticmethod
    def jinja_render_shape(shape: Shape, context: dict, loop_shape_ids: list):
        prev_shape = None
        for s in shape.child_shapes:  # type: Shape
            # manage for loops in template
            loop_shape_id = VisioFile.jinja_create_for_loop_if(s, prev_shape)
            if loop_shape_id:
                loop_shape_ids.append(loop_shape_id)
            prev_shape = s
            # manage 'set self' statements
            VisioFile.jinja_set_selfs(s, context)
            VisioFile.jinja_render_shape(shape=s, context=context, loop_shape_ids=loop_shape_ids)

    @staticmethod
    def jinja_set_selfs(shape: Shape, context: dict):
        # apply any {% self self.xxx = yyy %} statements in shape properties
        jinja_source = shape.text
        matches = re.findall(r'{% set self.(.*?)\s?=\s?(.*?) %}', jinja_source)  # non-greedy search for all {%...%} strings
        for m in matches:  # type: tuple  # expect ('property', 'value') such as ('x', '10') or ('y', 'n*2')
            property_name = m[0]
            value = "{{ "+m[1]+" }}"  # Jinja to be processed
            # todo: replace any self references in value with actual value - i.e. {% set self.x = self.x+1 %}
            self_refs = re.findall(r'self.(.*)[\s+-/*//]?', m[1])  # greedy search for all self.? between +, -, *, or /
            for self_ref in self_refs:  # type: tuple  # expect ('property', 'value') such as ('x', '10') or ('y', 'n*2')
                ref_val = str(shape.__getattribute__(self_ref[0]))
                value = value.replace('self.'+self_ref[0], ref_val)
            # use Jinja template to calculate any self refs found
            template = Template(value)  # value might be '{{ 1.0+2.4*3 }}'
            value = template.render(context)
            if property_name in ['x', 'y']:
                shape.__setattr__(property_name, value)

        # remove any {% set self %} statements, leaving any remaining text
        matches = re.findall('{% set self.*?%}', jinja_source)
        for m in matches:
            jinja_source = jinja_source.replace(m, '')  # remove Jinja 'set self' statement
        shape.text = jinja_source

    @staticmethod
    def unescape_jinja_statements(jinja_source):
        # unescape any text between {% ... %}
        jinja_source_out = jinja_source
        matches = re.findall('{%(.*?)%}', jinja_source)  # non-greedy search for all {%...%} strings
        for m in matches:
            unescaped = m.replace('&gt;', '>').replace('&lt;', '<')
            jinja_source_out = jinja_source_out.replace(m, unescaped)
        return jinja_source_out

    @staticmethod
    def jinja_create_for_loop_if(shape: Shape, previous_shape:Shape or None):
        # update a Shapes tag where text looks like a jinja {% for xxxx %} loop
        # move text to start of Shapes tag and add {% endfor %} at end of tag
        text = shape.text

        # use regex to find all loops
        jinja_loops = re.findall(r"{% for\s(.*?)\s%}", text)

        for loop in jinja_loops:
            jinja_loop_text = f"{{% for {loop} %}}"
            # move the for loop to start of shapes element (just before first Shape element)
            if previous_shape:
                if previous_shape.xml.tail:
                    previous_shape.xml.tail += jinja_loop_text
                else:
                    previous_shape.xml.tail = jinja_loop_text  # add jinja loop text after previous shape, before this element
            else:
                if shape.parent.xml.text:
                    shape.parent.xml.text += jinja_loop_text
                else:
                    shape.parent.xml.text = jinja_loop_text  # add jinja loop at start of parent, just before this element
            shape.text = shape.text.replace(jinja_loop_text, '')  # remove jinja loop from <Text> tag in element

            # add closing 'endfor' to just inside the shapes element, after last shape
            if shape.xml.tail:  # extend or set text at end of Shape element
                shape.xml.tail += "{% endfor %}"
            else:
                shape.xml.tail = '{% endfor %}'

        jinja_show_ifs = re.findall(r"{% showif\s(.*?)\s%}", text)  # find all showif statements
        # jinja_show_if - translate non-standard {% showif statement %} to valid jinja if statement
        for show_if in jinja_show_ifs:
            jinja_show_if = f"{{% if {show_if} %}}"  # translate to actual jinja if statement
            # move the for loop to start of shapes element (just before first Shape element)
            if previous_shape:
                previous_shape.xml.tail = str(previous_shape.xml.tail or '')+jinja_show_if  # add jinja loop text after previous shape, before this element
            else:
                shape.parent.xml.text = str(shape.parent.xml.text or '')+jinja_show_if  # add jinja loop at start of parent, just before this element

            # remove original jinja showif from <Text> tag in element
            shape.text = shape.text.replace(f"{{% showif {show_if} %}}", '')

            # add closing 'endfor' to just inside the shapes element, after last shape
            if shape.xml.tail:  # extend or set text at end of Shape element
                shape.xml.tail += "{% endif %}"
            else:
                shape.xml.tail = '{% endif %}'

        if jinja_loops:
            return shape.ID  # return shape ID if it is a loop, so that duplicate shape IDs can be updated

    @staticmethod
    def jinja_page_showif(page: Page, context: dict):
        text = page.name
        jinja_source = re.findall(r"{% showif\s(.*?)\s%}", text)
        if len(jinja_source):
            # process last matching value
            template_source = "{{ "+jinja_source[-1]+" }}"
            template = Template(template_source)  # value might be '{{ 1.0+2.4*3 }}'
            value = template.render(context)
            # is the value truthy - i.e. not 0, False, or empty string, tuple, list or dict
            print(f"jinja_page_showif(context={context}) statement: {template_source} returns: {type(value)} {value}")
            if value in ['False', '0', '', '()', '[]', '{}']:
                print("value in ['False', '0', '', '()', '[]', '{}']")
                return False  # page should be hidden
            # remove jinja statement from page name
            jinja_statement = re.match("{%.*?%}", page.name)[0]
            page.name = page.name.replace(jinja_statement, '')
        return True  # page should be left in

    @staticmethod
    def get_shape_id(shape: ET) -> str:
        return shape.attrib['ID']

    def increment_sub_shape_ids(self, shape: Shape, page, id_map: dict = None):
        id_map = self.increment_shape_ids(shape.xml, page, id_map)
        self.update_ids(shape.xml, id_map)
        for s in shape.child_shapes:
            id_map = self.increment_shape_ids(s.xml, page, id_map)
            self.update_ids(s.xml, id_map)
            if s.child_shapes:
                id_map = self.increment_sub_shape_ids(s, page, id_map)
        return id_map

    def copy_shape(self, shape: Element, page: Page) -> ET:
        """Insert shape into first Shapes tag in destination page, and return the copy.

        If destination page does not have a Shapes tag yet, create it.

        Parameters:
            shape (Element): The source shape to be copied. Use Shape.xml
            page (ElementTree): The page where the new Shape will be placed. Use Page.xml
            page_path (str): The filename of the page where the new Shape will be placed. Use Page.filename

        Returns:
            ElementTree: The new shape ElementTree

        """

        new_shape = ET.fromstring(ET.tostring(shape))

        page.set_max_ids()
        # find or create Shapes tag
        shapes_tag = page.xml.find(f"{namespace}Shapes")
        if shapes_tag is None:
            shapes_tag = Element(f"{namespace}Shapes")
            page.xml.getroot().append(shapes_tag)

        id_map = self.increment_shape_ids(new_shape, page) # page_obj)
        self.update_ids(new_shape, id_map)
        shapes_tag.append(new_shape)

        return new_shape

    def insert_shape(self, shape: Element, shapes: Element, page: ET, page_path: str) -> ET:
        # insert shape into shapes tag, and return updated shapes tag
        for page_obj in self.pages:
            if page_obj.filename == page_path:
                break

        id_map = self.increment_shape_ids(shape, page_obj)
        self.update_ids(shape, id_map)
        shapes.append(shape)
        return shapes

    def increment_shape_ids(self, shape: Element, page: Page, id_map: dict=None):
        if id_map is None:
            id_map = dict()
        self.set_new_id(shape, page, id_map)
        for e in shape.findall(f"{namespace}Shapes"):
            self.increment_shape_ids(e, page, id_map)
        for e in shape.findall(f"{namespace}Shape"):
            self.set_new_id(e, page, id_map)

        return id_map

    def set_new_id(self, element: Element, page: Page, id_map: dict):
        page.max_id += 1
        max_id = page.max_id
        if element.attrib.get('ID'):
            current_id = element.attrib['ID']
            id_map[current_id] = max_id  # record mappings
        element.attrib['ID'] = str(max_id)
        return max_id  # return new id for info

    def update_ids(self, shape: Element, id_map: dict):
        # update: <ns0:Cell F="Sheet.15! replacing 15 with new id using prepopulated id_map
        # cycle through shapes looking for Cell tag inside a Shape tag, which may be inside a Shapes tag
        for e in shape.findall(f"{namespace}Shapes"):
            self.update_ids(e, id_map)
        for e in shape.findall(f"{namespace}Shape"):
            # look for Cell elements
            cells = e.findall(f"{namespace}Cell[@F]")
            for cell in cells:
                f = str(cell.attrib['F'])
                if f.startswith("Sheet."):
                    # update sheet refs with new ids
                    shape_id = f.split('!')[0].split('.')[1]
                    new_id = id_map[shape_id]
                    new_f = f.replace(f"Sheet.{shape_id}",f"Sheet.{new_id}")
                    cell.attrib['F'] = new_f
        return shape

    def close_vsdx(self):
        try:
            # Remove extracted folder if there
            shutil.rmtree(self.directory)
        except (FileNotFoundError) as e:
            pass
        self.file_open = False

    def save_vsdx(self, new_filename=None):
        """save the VisioFile object as new vsdx file

        :param new_filename: path to save vsdx file
        :type new_filename: str

        """
        if not self.file_open:
            raise VisioFileNotOpen("Unable to save a file after being closed or outside of 'with' block.")
        # write pages.xml.rels
        xml_to_file(self.pages_xml_rels, f'{self.directory}/visio/pages/_rels/pages.xml.rels', self.zip_file_contents)

        # write pages.xml file - in case pages added removed
        xml_to_file(self.pages_xml, self._pages_filename(), self.zip_file_contents)

        # write the master pages to file
        for page in self.master_pages:  # type: Page
            xml_to_file(page.xml, page.filename, self.zip_file_contents)

        # write the pages to file
        for page in self.pages:  # type: Page
            xml_to_file(page.xml, page.filename, self.zip_file_contents)
            if page.rels_xml_filename:
                xml_to_file(page.rels_xml, page.rels_xml_filename, self.zip_file_contents)

        # write [content_Types].xml
        xml_to_file(self.content_types_xml, f'{self.directory}/[Content_Types].xml', self.zip_file_contents)

        # write app.xml
        if self.app_xml is not None:
            xml_to_file(self.app_xml, f'{self.directory}/docProps/app.xml', self.zip_file_contents)

        # write document.xml
        xml_to_file(self.document_xml, f'{self.directory}/visio/document.xml', self.zip_file_contents)

        # write document.xml.rels
        xml_to_file(self.document_xml_rels, f'{self.directory}/visio/_rels/document.xml.rels', self.zip_file_contents)

        # wrap up files into zip and rename to vsdx
        base_filename = self.filename[:-5]  # remove ".vsdx" from end
        if new_filename.find(os.sep) > 0:
            directory = new_filename[0:new_filename.rfind(os.sep)]
            if directory:
                if not os.path.exists(directory):
                    os.mkdir(directory)
        
        # write content from zip_file_contents to zip file directory
        if self.zip_file_contents:
            self._save_zip_file_contents_to_disk(new_filename or base_filename + '.zip')
            return

        shutil.make_archive(base_filename, 'zip', self.directory)

        if not new_filename:
            shutil.move(base_filename + '.zip', base_filename + '_new.vsdx')
        else:
            if new_filename[-5:] != '.vsdx':
                new_filename += '.vsdx'
            shutil.move(base_filename + '.zip', new_filename)
