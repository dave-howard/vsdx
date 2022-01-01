from __future__ import annotations
import shutil
import os

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from .pages import Page

import xml.etree.ElementTree as ET
from xml.etree.ElementTree import Element

from .shapes import Shape


class Connect:
    """Connect class to represent a connection between two `Shape` objects"""
    def __init__(self, xml: Element=None, page: Page=None):
        if page is None:
            return
        if type(xml) is Element:  # create from xml
            self.xml = xml
            self.page = page  # type: Page
            self.from_id = xml.attrib.get('FromSheet')  # ref to the connector shape
            self.to_id = xml.attrib.get('ToSheet')  # ref to the shape where the connector terminates
            self.from_rel = xml.attrib.get('FromCell')  # i.e. EndX / BeginX
            self.to_rel = xml.attrib.get('ToCell')  # i.e. PinX

    @staticmethod
    def create(page: Page=None, from_shape: Shape = None, to_shape: Shape = None) -> Shape:
        """Create a new Connect object between from_shape and to_shape

        :returns: a new Connect object
        :rtype: Shape
        """

        from .media import Media  # to break circular imports

        if from_shape and to_shape:  # create new connector shape and connect items between this and the two shapes
            # create new connect shape and get id
            media = Media()
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

            # update HeadingPairs and TitlesOfParts in app.xml
            if page.vis._get_app_xml_value('Masters') is None:
                page.vis._set_app_xml_value('Masters', '1')
            if 'Dynamic connector' not in page.vis._titles_of_parts_list():  # todo: replace static string with name from shape
                page.vis._add_titles_of_parts_item('Dynamic connector')

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
