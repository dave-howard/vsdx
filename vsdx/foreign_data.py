from __future__ import annotations
import io
import os

from xml.etree.ElementTree import Element

import vsdx

namespace = "{http://schemas.microsoft.com/office/visio/2012/main}"  # visio file name space
rel_namespace = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
relationship_namespace = "{http://schemas.openxmlformats.org/package/2006/relationships}"


class ForeignData:
    """ class to represent, and manipulate, the foreign data of a shape"""
    def __init__(self, xml: Element, page: vsdx.Page):
        self.xml = xml
        self.rels = list()
        self.page = page

        for rel in self.xml.findall(f"{namespace}Rel"):
            self.rels.append(Rel(parent=self, xml=rel))


class Rel:
    # TODO: parent can presumably be other stuff too?
    def __init__(self, parent: ForeignData, xml: Element):
        self.xml = xml
        self.parent = parent
        self.id = xml.attrib.get(f"{rel_namespace}id")

    # TODO: Presumably there are other types of rels that don't link to a file?
    @property
    def data_path(self) -> str or None:
        vis = self.parent.page.vis
        rel_targets = self.parent.page.rels_xml
        target_xml = rel_targets.find(f'{relationship_namespace}Relationship[@Id="{self.id}"]')
        if target_xml is None:
            return None
        target = target_xml.attrib.get("Target")
        if target is None:
            return None
        path = os.path.join(vis.pages_folder, target)
        return path

    @property
    def data(self) -> io.BytesIO:
        return self.parent.page.vis.zip_file_contents[self.data_path]
