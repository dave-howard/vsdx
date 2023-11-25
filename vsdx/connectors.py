from __future__ import annotations
import shutil
import os
import copy


import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

import vsdx
from .shapes import Shape


class Connect:
    """Connect class to represent a connection between two `Shape` objects"""
    def __init__(self, xml: Element=None, page: vsdx.Page=None):
        if page is None:
            return
        if type(xml) is Element:  # create from xml
            self.xml = xml
            self.page = page  # type: vsdx.Page
            self.from_id = xml.attrib.get('FromSheet')  # ref to the connector shape
            self.to_id = xml.attrib.get('ToSheet')  # ref to the shape where the connector terminates
            self.from_rel = xml.attrib.get('FromCell')  # i.e. EndX / BeginX
            self.to_rel = xml.attrib.get('ToCell')  # i.e. PinX

    @staticmethod
    def create(page: vsdx.Page=None, from_shape: Shape = None, to_shape: Shape = None) -> Shape:
        """Create a new Connect object between from_shape and to_shape

        :returns: a new Connect object
        :rtype: Shape
        """
        if from_shape and to_shape:  # create new connector shape and connect items between this and the two shapes
            # create new connect shape and get id
            media = vsdx.Media()
            connector_shape = media.straight_connector.copy(page)  # default to straight connector
            connector_shape.text = ''  # clear text used to find shape
            if not os.path.exists(page.vis._masters_folder):
                # Add masters folder to directory if not already present
                shutil.copytree(media._media_vsdx._masters_folder, page.vis._masters_folder)
                page.vis.load_master_pages()  # load copied master page files into VisioFile object
                # add new master to document relationship
                page.vis._add_document_rel(rel_type="http://schemas.microsoft.com/visio/2010/relationships/masters",
                                           target="masters/masters.xml")
                # create masters/master1 elements in [Content_Types].xml
                page.vis._add_content_types_override(content_type="application/vnd.ms-visio.masters+xml",
                                                     part_name_path="/visio/masters/masters.xml")
                page.vis._add_content_types_override(content_type="application/vnd.ms-visio.master+xml",
                                                     part_name_path="/visio/masters/master1.xml")
                # create an initial copy of page_rels from media and attach to this page
                page_rels_xml = copy.deepcopy(media.rels_xml)
                page.rels_xml = page_rels_xml
            elif connector_shape.shape_name not in page.vis._titles_of_parts_list():
                print(f"Warning: Updating existing Page/Master relationships not yet fully implemented. "
                      f"This may cause unexpected outputs.")
                # vsdx has masters - but not this shape
                # todo: Complete this scenario
                #print("conn master page", connector_shape.master_shape.page.filename)
                #print("max page file num", [p.filename[-5:-4] for p in page.vis.master_pages])
                #print("max page id", [p.page_id for p in page.vis.master_pages])
                rel_num = max([int(p.filename[-5:-4]) for p in page.vis.master_pages]) +1
                master_file_path = os.path.join(page.vis.directory, 'visio', 'masters', f'master{rel_num}.xml')
                #print(f"m_num={rel_num} master_file_path={master_file_path}")
                shutil.copy(connector_shape.master_shape.page.filename, master_file_path)
                # todo: ensure master page ID and RId is unique, update shape master_id to refer to new master
                # todo: update mast file name, and add content type override
                # todo: update masters.xml file contents?
                # todo: update visio/pages/_rels/page3.xml.rels - add: <Relationship Id="rId3" Type="http://schemas.microsoft.com/visio/2010/relationships/master" Target="../masters/master1.xml"/>
                rels = page.rels_xml.getroot()
                new_rel = ET.fromstring(f'<Relationship  xmlns="{vsdx.document_rels_namespace[1:-1]}" '
                                        f'Type="http://schemas.microsoft.com/visio/2010/relationships/master" />')
                new_rel.attrib['Id'] = f"rID{rel_num}"
                new_rel.attrib['Target'] = f"../masters/master{rel_num}.xml"
                rels.append(new_rel)
                page.vis._add_content_types_override(content_type="application/vnd.ms-visio.master+xml",
                                                     part_name_path=f"/visio/masters/master{rel_num}.xml")
            else:
                # vsdx has this master shape, but not related to this page
                master_page = page.vis.master_index.get(connector_shape.shape_name)  # type: vsdx.Page
                rel_num = int(master_page.rel_id[-1])
                rels = page.rels_xml.getroot()
                new_rel = ET.fromstring(f'<Relationship  xmlns="{vsdx.document_rels_namespace[1:-1]}" '
                                        f'Type="http://schemas.microsoft.com/visio/2010/relationships/master" />')
                new_rel.attrib['Id'] = master_page.rel_id
                new_rel.attrib['Target'] = "../masters/master1.xml"
                rels.append(new_rel)

            # update HeadingPairs and TitlesOfParts in app.xml
            if page.vis._get_app_xml_value('Masters') is None:
                page.vis._set_app_xml_value('Masters', '1')

            if connector_shape.shape_name not in page.vis._titles_of_parts_list():  # todo: replace static string with name from shape
                page.vis._add_titles_of_parts_item(connector_shape.shape_name)

            # copy style used by new connector shape
            if not page.vis._get_style_by_id(connector_shape.master_shape.line_style_id):
                # assume same if is ok, todo: use names for match and increment IDs
                media_style = media._media_vsdx._get_style_by_id(connector_shape.master_shape.line_style_id)
                page.vis._style_sheets().append(media_style)
            media._media_vsdx.close_vsdx()

            # set Begin and End Trigger formulae for the new shape - linking to shapes in destination page
            beg_trigger = connector_shape.cells.get('BegTrigger')
            beg_trigger.formula = beg_trigger.formula.replace('Sheet.1!', f'Sheet{from_shape.ID}!')
            end_trigger = connector_shape.cells.get('EndTrigger')
            end_trigger.formula = end_trigger.formula.replace('Sheet.2!', f'Sheet{to_shape.ID}!')

            # create connect relationships
            # todo: FromPart="12" and ToPart="3" represent the part of a shape to connection is from/to
            end_connect_xml = f'<Connect xmlns="http://schemas.microsoft.com/office/visio/2012/main" FromSheet="{connector_shape.ID}" FromCell="EndX" FromPart="12" ToSheet="{to_shape.ID}" ToCell="PinX" ToPart="3"/>'
            beg_connect_xml = f'<Connect xmlns="http://schemas.microsoft.com/office/visio/2012/main" FromSheet="{connector_shape.ID}" FromCell="BeginX" FromPart="9" ToSheet="{from_shape.ID}" ToCell="PinX" ToPart="3"/>'

            # Add these new connection relationships to the page
            page.add_connect(Connect(xml=ET.fromstring(end_connect_xml), page=page))
            page.add_connect(Connect(xml=ET.fromstring(beg_connect_xml), page=page))
            #print(vsdx.pretty_print_element(connector_shape.xml))
            #print(connector_shape.geometry)

            connector_shape.set_start_and_finish(from_shape.center_x_y, to_shape.center_x_y)
            print([(m.rel_id, m.page_id) for m in page.vis.master_pages])
            return connector_shape

    @property
    def shape_id(self):
        # ref to the shape where the connector terminates - convenience property
        return self.to_id

    @property
    def shape(self) -> Shape:
        return self.page.find_shape_by_id(self.shape_id)

    @property
    def connector_shape_id(self):
        # ref to the connector shape - convenience property
        return self.from_id

    @property
    def connector_shape(self) -> Shape:
        return self.page.find_shape_by_id(self.connector_shape_id)

    def __repr__(self):
        return f"Connect: from={self.from_id} to={self.to_id} connector_id={self.connector_shape_id} shape_id={self.shape_id}"
