import pytest
from datetime import datetime
import os

from typing import List

from vsdx import Connect
from vsdx import Page
from vsdx import Shape
from vsdx import VisioFile

# code to get basedir of this test file in either linux/windows
basedir = os.path.dirname(os.path.relpath(__file__))


@pytest.mark.parametrize("filename, count", [("test1.vsdx", 1), ("test2.vsdx", 1)])
def test_get_page_shapes(filename: str, count: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        print(f"shape count={len(page.shapes)}")
        assert len(page.shapes) == count


@pytest.mark.parametrize("filename, count", [("test1.vsdx", 4), ("test2.vsdx", 6)])
def test_get_page_sub_shapes(filename: str, count: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        shapes = page.sub_shapes()
        print(f"shape count={len(shapes)}")
        assert len(shapes) == count


@pytest.mark.parametrize("filename, shape_id", [("test1.vsdx", "6"), ("test2.vsdx", "6")])
def test_get_shape_with_text(filename: str, shape_id: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        shape = page.find_shape_by_text('{{date}}')  # type: Shape
        assert shape.ID == shape_id


@pytest.mark.parametrize("filename", ["test1.vsdx", "test2.vsdx", "test3_house.vsdx"])
def test_apply_context(filename: str):
    date_str = str(datetime.today().date())
    context = {'scenario': 'test',
               'date': date_str}
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_VISfilter_applied.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        original_shape = page.find_shape_by_text('{{date}}')  # type: Shape
        assert original_shape.ID
        page.apply_text_context(context)
        vis.save_vsdx(out_file)

    # open and find date_str
    with VisioFile(out_file) as vis:
        page = vis.get_page(0)  # type: Page
        updated_shape = page.find_shape_by_text(date_str)  # type: Shape
        assert updated_shape.ID == original_shape.ID


@pytest.mark.parametrize("filename", ["test1.vsdx", "test2.vsdx", "test3_house.vsdx"])
def test_find_replace(filename: str):
    old = 'Shape'
    new = 'Figure'
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_VISfind_replace_applied.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        original_shapes = page.find_shapes_by_text(old)  # type: List[Shape]
        shape_ids = [s.ID for s in original_shapes]
        page.find_replace(old, new)
        vis.save_vsdx(out_file)

    # open and find date_str
    with VisioFile(out_file) as vis:
        page = vis.get_page(0)  # type: Page
        # test that each shape if has 'new' str in text
        for shape_id in shape_ids:
            shape = page.find_shape_by_id(shape_id)
            assert new in shape.text


# DataProperty

@pytest.mark.parametrize(("filename", "page_index", "expected_shape_name", "property_label"),
                         [("test1.vsdx", 0, "Shape Text", "my_property_label"),
                          ("test6_shape_properties.vsdx", 0, "Shape One", "my_property_label"),
                          ("test6_shape_properties.vsdx", 0, "Shape One", "my_second_property_label"),
                          ("test6_shape_properties.vsdx", 0, "Shape Three", "my_third_property_label"),
                          ("test6_shape_properties.vsdx", 2, "A", "master_Prop"),
                          ])
def test_find_shape_by_data_property_label(filename: str, page_index: int, expected_shape_name: str, property_label: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shape = vis.pages[page_index].find_shape_by_property_label(property_label)  # type: Shape

        # test that a shape is returned
        assert shape

        # test that returned shape has a DataProperty with property_name
        assert shape.data_properties[property_label]

        # test iot's the expected shape
        assert shape.text.replace('\n', '') == expected_shape_name


@pytest.mark.parametrize(("filename", "page_index", "expected_shape_names", "property_label"),
                         [("test1.vsdx", 0, ["Shape Text"], "my_property_label"),
                          ("test6_shape_properties.vsdx", 0, ["Shape One", "Shape Two"], "my_property_label"),
                          ("test6_shape_properties.vsdx", 0, ["Shape One", "Shape Three"], "my_second_property_label"),
                          ("test6_shape_properties.vsdx", 1, ["Shape A", "Shape C"], "label_one"),
                          ("test6_shape_properties.vsdx", 2, ["A", "B", "C"], "master_Prop"),
                          ])
def test_find_shapes_by_data_property_label(filename: str, page_index: int, expected_shape_names: list, property_label: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shapes = vis.pages[page_index].find_shapes_by_property_label(property_label)

        # test that a shape is returned
        assert len(shapes) == len(expected_shape_names)

        # get list of shape names (text) without carriage returns
        shape_names = sorted([s.text.replace('\n', '') for s in shapes])
        print(f"shape names={shape_names}")

        # test that the shapes returned have the expected names
        assert shape_names == expected_shape_names

        # test that each returned shape has a DataProperty with property_name
        for s in shapes:
            assert s.data_properties[property_label]


# Connectors

@pytest.mark.parametrize(("filename", "expected_connects"),
                         [('test4_connectors.vsdx', ["from 7 to 5", "from 7 to 2", "from 6 to 2", "from 6 to 1"]),
                          ])
def test_find_page_connects(filename: str, expected_connects: list):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: Page
        actual_connects = list()
        for c in page.connects:  # type: Connect
            actual_connects.append(f"from {c.from_id} to {c.to_id}")
        assert sorted(actual_connects) == sorted(expected_connects)


@pytest.mark.parametrize(("filename", "shape_a_id", "shape_b_id", "expected_connector_ids"),
                         [
                             ('test4_connectors.vsdx', "1", "2", ["6"]),
                             ('test4_connectors.vsdx', "1", "5", []),  # expect no connections
                          ])
def test_find_connectors_between_ids(filename: str, shape_a_id: str, shape_b_id: str,  expected_connector_ids: list):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: Page
        connectors = page.get_connectors_between(shape_a_id=shape_a_id, shape_b_id=shape_b_id)
        actual_connector_ids = sorted([c.ID for c in connectors])
        assert sorted(expected_connector_ids) == list(actual_connector_ids)


@pytest.mark.parametrize(("filename", "shape_a_text", "shape_b_text", "expected_connector_ids"),
                         [
                             ('test4_connectors.vsdx', "Shape A", "Shape B", ["6"]),
                             ('test4_connectors.vsdx', "Shape A", "Shape C", []),  # expect no connections
                          ])
def test_find_connectors_between_shapes(filename: str, shape_a_text: str, shape_b_text: str,  expected_connector_ids: list):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: Page
        connectors = page.get_connectors_between(shape_a_text=shape_a_text, shape_b_text=shape_b_text)
        actual_connector_ids = sorted([c.ID for c in connectors])
        assert sorted(expected_connector_ids) == list(actual_connector_ids)


@pytest.mark.parametrize(("filename", "shape_a_text", "shape_b_text"),
                         [
                             ('test7_with_connector.vsdx', "Shape Text", "Shape to remove"),
                             ('test1.vsdx', "Shape to remove", "Shape Text"),
                             ('test1.vsdx', "Shape Text", "Shape to copy"),
                             ('test1.vsdx', "Shape to copy", "Shape to remove"),
                             ('test4_connectors.vsdx', "Shape A", "Shape C"),
                             ('test4_connectors.vsdx', "Shape C", "Shape B"),
                          ])
def test_add_connect_between_shapes(filename: str, shape_a_text: str, shape_b_text: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_add_connect_between_'+shape_a_text+'_'+shape_b_text+'.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        print(f"filename:{out_file}")
        page = vis.pages[0]  # type: Page
        from_shape = page.find_shape_by_text(shape_a_text)
        to_shape = page.find_shape_by_text(shape_b_text)
        c = Connect.create(page=page, from_shape=from_shape, to_shape=to_shape)
        print(f"before move() x,y={c.x},{c.y} geo:{c.geometry}")
        c.move(from_shape.x - c.x, from_shape.y - c.y)
        print(f"after move() x,y={c.x},{c.y} geo:{c.geometry}")
        c.geometry.set_line_to(to_shape.x, to_shape.y)
        print(f"after set_line_to() x,y={c.x},{c.y} geo:{c.geometry}")
        new_connector_id = c.ID
        #conns_shown = []
        #for conn in page.connects:
        #    if conn.connector_shape.ID not in conns_shown:
        #        print(f"conn between {[s.ID for s in conn.connector_shape.connected_shapes]} {conn.from_id}->{conn.to_id}:{VisioFile.pretty_print_element(conn.connector_shape.xml)}")
        #        conns_shown.append(conn.connector_shape.ID)
        #    print(VisioFile.pretty_print_element(conn.xml))
        vis.save_vsdx(out_file)

        # re-open saved file and check it is changed as expected
        with VisioFile(out_file) as vis:
            page = vis.pages[0]
            connector_ids = [c.connector_shape_id for c in page.connects]
            # new shape is referenced as connector_shape for new connector relationship
            assert new_connector_id in connector_ids
            # new shape exists in page
            c = page.find_shape_by_id(new_connector_id)
            #print(f"conn_shape.line_to_x={c.line_to_x}")
            assert page.find_shape_by_id(new_connector_id)
