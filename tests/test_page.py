import pytest
from datetime import datetime
import os

from typing import List

import vsdx
from vsdx import Connect
from vsdx import Page
from vsdx import Shape
from vsdx import VisioFile
from vsdx import Cell

# code to get basedir of this test file in either linux/windows
basedir = os.path.dirname(os.path.relpath(__file__))


@pytest.mark.parametrize("filename, count", [("test1.vsdx", 1), ("test2.vsdx", 1)])
def test_get_page_shapes(filename: str, count: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        print(f"shape count={len(page.shapes)}")
        assert len(page.shapes) == count


@pytest.mark.parametrize("filename, page_index, height_width",
                         [("test1.vsdx", 0, (8.26771653543307, 11.69291338582677)),
                          ("test2.vsdx", 0, (8.26771653543307, 11.69291338582677)),
                          ("test1.vsdx", 1, (8.26771653543307, 11.69291338582677)),
                          ("test2.vsdx", 1, (8.26771653543307, 11.69291338582677)),])
def test_get_page_size(filename: str, page_index: int, height_width: tuple):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index]
        print(VisioFile.pretty_print_element(page._pagesheet_xml))
        print(f"\n w x h={page.width} x {page.height}")
        assert (page.width, page.height) == height_width


@pytest.mark.parametrize("filename, page_index",
                         [("test1.vsdx", 0),
                          ("test2.vsdx", 0),
                          ("test1.vsdx", 1),
                          ("test2.vsdx", 1),])
def test_set_page_size(filename: str, page_index: int):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_set_page_size_{page_index}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index]
        #print(VisioFile.pretty_print_element(page._pagesheet_xml))
        print(f"\n w x h={page.width} x {page.height}")
        page.width = page.width // 1
        page.height = page.height // 1
        vis.save_vsdx(out_file)

        with VisioFile(out_file) as vis:
            page = vis.pages[page_index]
            print(f"\n w x h={page.width} x {page.height}")
            assert page.width == page.width // 1
            assert page.height == page.height // 1



@pytest.mark.parametrize("filename, page_index, page_name",
                         [("test1.vsdx", 0, "Page-1"),
                          ("test2.vsdx", 0, "Page-1"),
                          ("test1.vsdx", 1, "Page-2"),
                          ("test2.vsdx", 1, "Page-2"),])
def test_get_page_name(filename: str, page_index: int, page_name: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index]
        print(VisioFile.pretty_print_element(page._pagesheet_xml))
        assert page.name == page_name


@pytest.mark.parametrize("filename, page_index, page_name",
                         [("test1.vsdx", 0, "Page1"),
                          ("test2.vsdx", 0, "Page1"),
                          ("test1.vsdx", 1, "Page2"),
                          ("test2.vsdx", 1, "Page2"),])
def test_set_page_name(filename: str, page_index: int, page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_set_page_name_{page_name}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index]
        print(VisioFile.pretty_print_element(page._pagesheet_xml))
        page.page_name = page_name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.pages[page_index]
        assert page.page_name == page_name


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
                             ('test8_simple_connector.vsdx', "Shape A", "Shape B"),
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
        new_connector_id = c.ID

        c.text = "NEW"
        vis.save_vsdx(out_file)

        # re-open saved file and check it is changed as expected
        with VisioFile(out_file) as vis:
            page = vis.pages[0]
            connector_ids = [c.connector_shape_id for c in page.connects]
            # new shape is referenced as connector_shape for new connector relationship
            assert new_connector_id in connector_ids
            # new shape exists in page
            c = page.find_shape_by_id(new_connector_id)
            assert page.find_shape_by_id(new_connector_id)


def fl(v: float):
    if type(v) is float:
        return f"{v:.2g}"


def add_shape_info(s: Shape):
    s.text = s.text + f" x,y={fl(s.x)}, {fl(s.y)}"
    if s.begin_x is not None:
        s.text = s.text + f" bx,by={fl(s.begin_x)}, {fl(s.begin_y)}"
        s.text = s.text + f" ex,ey={fl(s.end_x)}, {fl(s.end_y)}"
    s.text = s.text + f" w,h={fl(s.width)}, {fl(s.height)}"
    cx, cy = s.center_x_y
    s.text = s.text + f" cx,cy={fl(cx)}, {fl(cy)}"
    s.text = s.text + f" locx,locy={fl(s.loc_x)}, {fl(s.loc_y)}"


@pytest.mark.parametrize(("filename", "shape_text", "lx", "ly"),
                         [
                             ('test9_rect_and_line.vsdx', "Rect A", 2.0, 8.0),
                             ('test9_rect_and_line.vsdx', "Line A", 2.0, 8.0),
                             ('test9_rect_and_line.vsdx', "Conn A", 2.0, 8.0),
                             ('test9_rect_and_line.vsdx', "Rect A", 2.0, 5.0),
                             ('test9_rect_and_line.vsdx', "Line A", 2.0, 5.0),
                             ('test9_rect_and_line.vsdx', "Conn A", 2.0, 5.0),
                          ])
def test_copy_and_move_shape(filename: str, shape_text: str, lx: float, ly: float):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_copy_and_move_shape_{shape_text}_{lx}_{ly}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        print(f"filename:{out_file}")
        page = vis.pages[0]  # type: Page
        shape = page.find_shape_by_text(shape_text)
        print(vsdx.pretty_print_element(shape.xml))
        # should work for shapes
        cp1 = shape.copy()
        add_shape_info(shape)
        cp1.x, cp1.y = lx, ly
        cp1.text = cp1.text + f"lx,ly={fl(lx)}, {fl(ly)}"
        if shape.begin_x is not None:
            # should work for lines
            cp1.begin_x, cp1.begin_y = lx, ly
            cp1.end_x, cp1.end_y = lx + cp1.width, ly + cp1.height
            cp1.x, cp1.y = cp1.center_x_y
        add_shape_info(cp1)

        vis.save_vsdx(out_file)


@pytest.mark.parametrize(("filename", "shape_text", "start", "finish"),
                         [
                             ('test9_rect_and_line.vsdx', "Line A", (2.0, 7.0), (3.0, 8.0)),
                             ('test9_rect_and_line.vsdx', "Conn A", (2.0, 7.0), (3.0, 8.0)),
                          ])
def test_copy_and_move_line(filename: str, shape_text: str, start: tuple, finish: tuple):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_copy_and_move_line_{shape_text}_{start}_{finish}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        print(f"filename:{out_file}")
        page = vis.pages[0]  # type: Page
        shape = page.find_shape_by_text(shape_text)
        print(vsdx.pretty_print_element(shape.xml))
        # should work for shapes
        cp1 = shape.copy()
        add_shape_info(shape)
        cp1.x, cp1.y = start
        cp1.text = cp1.text + f"lx,ly={fl(cp1.x)}, {fl(cp1.y)}"
        if cp1.begin_x is not None:
            is_connector = cp1.shape_name == 'Dynamic connector'
            # should work for lines
            cp1.begin_x, cp1.begin_y = start
            cp1.end_x, cp1.end_y = finish
            cp1.width = cp1.end_x - cp1.begin_x
            if is_connector:
                cp1.height = cp1.end_y - cp1.begin_y  # connector has a height
            else:
                cp1.height = 0.0  # line height is always zero
            cp1.x, cp1.y = start # cp1.center_x_y
            cp1.geometry.set_move_to(0.0, 0.0)
            cp1.geometry.set_line_to(cp1.width, cp1.height)
            txt_pin_x = cp1.cells.get('TxtPinX')
            txt_pin_y = cp1.cells.get('TxtPinY')
            if txt_pin_x and txt_pin_y:
                if is_connector:
                    txt_pin_x.value, txt_pin_y.value = cp1.width/2, cp1.height/2
                else:
                    txt_pin_x.value, txt_pin_y.value = cp1.center_x_y
                cp1.set_cell_value(name='Control/TextPosition/X', value=txt_pin_x.value)
                cp1.set_cell_value(name='Control/TextPosition/Y', value=txt_pin_y.value)
                cp1.set_cell_value(name='Control/TextPosition/XDyn', value=txt_pin_x.value)
                cp1.set_cell_value(name='Control/TextPosition/YDyn', value=txt_pin_y.value)
                #print(cp1.cells.keys())
            cells = list(cp1.cells.values())+cp1.geometry.cells
            for r in cp1.geometry.rows.values():
                cells.extend(r.cells.values())
            #print(cells)
            for c in cells:  # type: Cell
                v = None
                formula = c.formula
                if formula:
                    if formula == 'Inh' and cp1.master_shape:
                        print(f'Inh: {c.name} {cp1.master_shape.cells.get(c.name)}')
                        master_c = cp1.master_shape.cells.get(c.name)
                        formula = master_c.formula if master_c else formula
                    v = vsdx.calc_value(cp1, formula)
                    if v is not None:
                        c.value = v
                #print(f"c={c.name} f={c.formula} v={v}")

            print(vsdx.pretty_print_element(cp1.xml))
        add_shape_info(cp1)

        vis.save_vsdx(out_file)