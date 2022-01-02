import pytest
from vsdx import VisioFile, namespace, vt_namespace, ext_prop_namespace, PagePosition, Media
from datetime import datetime
import os
from typing import List

# code to get basedir of this test file in either linux/windows
filename = str(__file__).split(os.sep)[-1]  # get filename as str after any '/'
basedir = os.path.relpath(__file__)[:-(len(filename))]  # remove filename to get directory


def test_invalid_file_type():
    filename = __file__
    print(f"Opening invalid but existing {filename}")
    try:
        with VisioFile(filename):
            assert False  # don't expect to get here
    except TypeError as e:
        print(e)  # expect to get here


def test_open_rel_path():
    # test opening media file (not in tests directory)with absolute path
    print("__file__:" + __file__)
    print("os.path.relpath(__file__):" + os.path.relpath(__file__))
    print("basedir:" + basedir)
    filename = os.path.join(basedir, 'test1.vsdx')
    print("filename:"+filename)
    assert(os.path.exists(filename))
    with VisioFile(filename) as vis:
        for p in vis.pages:
            print(p.name, p.shapes)


def test_open_abs_path():
    # test opening media file (not in tests directory)with absolute path
    media_file_path = Media()._media_vsdx.filename
    filename = os.path.abspath(media_file_path)

    assert(os.path.exists(filename))
    with VisioFile(filename) as vis:
        for p in vis.pages:
            print(p.name, p.shapes)


def test_open_abs_path_save_rel_path():
    # test opening media file (not in tests directory)with absolute path
    media_file_path = Media()._media_vsdx.filename
    filename = os.path.abspath(media_file_path)

    assert(os.path.exists(filename))
    output_file = os.path.join(basedir, 'out', 'abs_to_rel_out.vsdx')
    with VisioFile(filename) as vis:
        vis.save_vsdx(output_file)


def test_open_abs_path_save_abs_path():
    # test opening media file (not in tests directory)with absolute path
    media_file_path = Media()._media_vsdx.filename
    filename = os.path.abspath(media_file_path)

    assert(os.path.exists(filename))
    output_file = os.path.abspath(os.path.join(basedir, 'out', 'abs_to_abs_out.vsdx'))
    print('output_file', output_file)
    with VisioFile(filename) as vis:
        vis.save_vsdx(output_file)


def test_file_closure():
    filename = os.path.join(basedir, 'test1.vsdx')
    directory = f"./{filename.rsplit('.', 1)[0]}"
    with VisioFile(filename):
        # confirm directory exists
        assert os.path.exists(directory)
    # confirm directory is gone
    assert not os.path.exists(directory)


@pytest.mark.parametrize("filename, page_name", [("test1.vsdx", "Page-1"), ("test2.vsdx", "Page-1")])
def test_get_page(filename: str, page_name: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        # confirm page name as expected
        assert page.name == page_name


@pytest.mark.parametrize("filename, count", [("test1.vsdx", 1), ("test2.vsdx", 1)])
def test_get_page_shapes(filename: str, count: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        print(f"shape count={len(page.shapes)}")
        assert len(page.shapes) == count


@pytest.mark.parametrize("filename, count", [("test1.vsdx", 4), ("test2.vsdx", 6)])
def test_get_page_sub_shapes(filename: str, count: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        shapes = page.sub_shapes()
        print(f"shape count={len(shapes)}")
        assert len(shapes) == count


@pytest.mark.parametrize("filename, expected_locations",
                         [("test1.vsdx", "1.33,10.66 4.13,10.66 6.94,10.66 2.33,9.02 "),
                          ("test2.vsdx", "2.33,8.72 1.33,10.66 4.13,10.66 5.91,8.72 1.61,8.58 3.25,8.65 ")])
def test_shape_locations(filename: str, expected_locations: str):
    print("=== list_shape_locations ===")
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        shapes = page.sub_shapes()
        locations = ""
        for s in shapes:  # type: VisioFile.Shape
            locations += f"{s.x:.2f},{s.y:.2f} "
        print(f"Expected:{expected_locations}")
        print(f"  Actual:{locations}")
    assert locations == expected_locations

@pytest.mark.skip
@pytest.mark.parametrize("filename, shape_id, expected_center",
                         [("test1.vsdx", "1", (1,1)),
                          ("test1.vsdx", "2", (1,1)),
                          ("test1.vsdx", "5", (1,1)),
                          ("test1.vsdx", "6", (1,1))])
def test_shape_center(filename: str, shape_id: str, expected_center: str):

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        print(f"Shape ids: {[s.ID for s in page.sub_shapes()]}")
        shape = page.find_shape_by_id(shape_id)

        print(f"shape {shape.ID} center={shape.center_x_y}")
        print(f"x {shape.x} y={shape.y}")
        print(f"end_x {shape.end_x} end_y={shape.end_y}")
        assert shape.center_x_y == expected_center


@pytest.mark.parametrize("filename, shape_id", [("test1.vsdx", "6"), ("test2.vsdx", "6")])
def test_get_shape_with_text(filename: str, shape_id: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        shape = page.find_shape_by_text('{{date}}')  # type: VisioFile.Shape
        assert shape.ID == shape_id


@pytest.mark.parametrize("filename", ["test1.vsdx", "test2.vsdx", "test3_house.vsdx"])
def test_apply_context(filename: str):
    date_str = str(datetime.today().date())
    context = {'scenario': 'test',
               'date': date_str}
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_VISfilter_applied.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        original_shape = page.find_shape_by_text('{{date}}')  # type: VisioFile.Shape
        assert original_shape.ID
        page.apply_text_context(context)
        vis.save_vsdx(out_file)

    # open and find date_str
    with VisioFile(out_file) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        updated_shape = page.find_shape_by_text(date_str)  # type: VisioFile.Shape
        assert updated_shape.ID == original_shape.ID


@pytest.mark.parametrize("filename", ["test1.vsdx", "test2.vsdx", "test3_house.vsdx"])
def test_find_replace(filename: str):
    old = 'Shape'
    new = 'Figure'
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_VISfind_replace_applied.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        original_shapes = page.find_shapes_by_text(old)  # type: List[VisioFile.Shape]
        shape_ids = [s.ID for s in original_shapes]
        page.find_replace(old, new)
        vis.save_vsdx(out_file)

    # open and find date_str
    with VisioFile(out_file) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        # test that each shape if has 'new' str in text
        for shape_id in shape_ids:
            shape = page.find_shape_by_id(shape_id)
            assert new in shape.text


@pytest.mark.parametrize("filename", ["test2.vsdx", "test3_house.vsdx"])
def test_remove_shape(filename: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_shape_removed.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shapes = vis.pages[0].shapes
        # get shape to remove
        s = shapes[0].find_shape_by_text('Shape to remove')  # type: VisioFile.Shape
        assert s  # check shape found
        s.remove()
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        shapes = vis.pages[0].shapes
        # get shape that should have been removed
        s = shapes[0].find_shape_by_text('Shape to remove')  # type: VisioFile.Shape
        assert s is None  # check shape not found


@pytest.mark.parametrize(("filename", "shape_names", "shape_locations"),
                         [("test1.vsdx", {"Shape to remove"}, {(1.0, 1.0)})])
def test_set_shape_location(filename: str, shape_names: set, shape_locations: set):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_set_shape_location.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shapes = vis.pages[0].shapes
        # move shapes in list
        for (shape_name, x_y) in zip(shape_names, shape_locations):
            s = shapes[0].find_shape_by_text(shape_name)  # type: VisioFile.Shape
            assert s  # check shape found
            assert s.x
            assert s.y
            print(f"Moving shape '{shape_name}' from {s.x}, {s.y} to {x_y}")
            s.x = x_y[0]
            s.y = x_y[1]
        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        shapes = vis.pages[0].shapes
        # get each shape that should have been moved
        for (shape_name, x_y) in zip(shape_names, shape_locations):
            s = shapes[0].find_shape_by_text(shape_name)  # type: VisioFile.Shape
            assert s  # check shape found
            assert s.x == x_y[0]
            assert s.y == x_y[1]


@pytest.mark.parametrize(("filename", "shape_name"),
                         [("test1.vsdx", "Shape Text",),
                          ])
def test_get_shape_data_properties_type_is_dict_of_data_property(filename: str, shape_name: str):
    # check that Shape.data_properties is a dict of ShapeProperty
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shape = vis.pages[0].find_shape_by_text(shape_name)  # type: VisioFile.Shape

        props = shape.data_properties

        # check overall type
        assert type(props) is dict

        # check each item type
        for prop in props.values():
            assert type(prop) is VisioFile.DataProperty


@pytest.mark.parametrize(("filename", "page_index", "shape_name", "property_dict"),
                         [("test1.vsdx", 0, "Shape Text", {"my_property_label": "property value",
                                                        "my_second_property_label": "another value"}),
                          ("test6_shape_properties.vsdx", 2, "A", {"master_Prop": "master prop value"}),
                          ("test6_shape_properties.vsdx", 2, "B", {"master_Prop": "master prop value",
                                                                   "shape_prop": "shape property value"}),
                          ("test6_shape_properties.vsdx", 2, "C", {"master_Prop": "override"}),
                          ])
def test_get_shape_data_properties(filename: str, page_index: int, shape_name: str, property_dict: dict):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shape = vis.pages[page_index].find_shape_by_text(shape_name)  # type: VisioFile.Shape

        props = shape.data_properties

        # check lengths are same
        assert len(props) == len(property_dict)

        # check each key/value is same as expected
        for property_label in props.keys():
            prop = props[property_label]
            print(f"prop: lbl:'{prop.label}' name:'{prop.name}': val:'{prop.value}'")
            assert prop.value == property_dict.get(property_label)


@pytest.mark.parametrize(("filename", "page_index", "expected_shape_name", "property_label"),
                         [("test1.vsdx", 0, "Shape Text", "my_property_label"),
                          ("test6_shape_properties.vsdx", 0, "Shape One", "my_property_label"),
                          ("test6_shape_properties.vsdx", 0, "Shape One", "my_second_property_label"),
                          ("test6_shape_properties.vsdx", 0, "Shape Three", "my_third_property_label"),
                          ("test6_shape_properties.vsdx", 2, "A", "master_Prop"),
                          ])
def test_find_shape_by_data_property_label(filename: str, page_index: int, expected_shape_name: str, property_label: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shape = vis.pages[page_index].find_shape_by_property_label(property_label)  # type: VisioFile.Shape

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


@pytest.mark.parametrize(("filename", "page_index", "container_shape_name", "expected_shape_name", "property_label"),
                         [("test6_shape_properties.vsdx", 1, "Container", "Shape A", "label_one"),
                          ])
def test_find_sub_shape_by_data_property_label(filename: str, page_index: int, container_shape_name: str, expected_shape_name: list, property_label: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        container_shape = vis.pages[page_index].find_shape_by_text(container_shape_name)

        shape = container_shape.find_shape_by_property_label(property_label)

        # test that a shape is returned
        assert shape

        # test that the shape returned has the expected names
        assert shape.text.replace('\n', '') == expected_shape_name


@pytest.mark.parametrize(("filename", "shape_names", "shape_x_y_deltas"),
                         [("test1.vsdx", {"Shape to remove"}, {(1.0, 1.0)}),
                          ("test4_connectors.vsdx", {"Shape B"}, {(1.0, 1.0)})])
def test_move_shape(filename: str, shape_names: set, shape_x_y_deltas: set):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_move_shape.vsdx')
    expected_shape_locations = dict()

    with VisioFile(os.path.join(basedir, filename)) as vis:
        shapes = vis.pages[0].shapes
        # move shapes in list
        for (shape_name, x_y) in zip(shape_names, shape_x_y_deltas):
            s = shapes[0].find_shape_by_text(shape_name)  # type: VisioFile.Shape
            assert s  # check shape found
            assert s.x
            assert s.y
            expected_shape_locations[shape_name] = (s.x + x_y[0], s.y + x_y[1])
            s.move(x_y[0], x_y[1])

        vis.save_vsdx(out_file)
    print(f"expected_shape_locations={expected_shape_locations}")

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        shapes = vis.pages[0].shapes
        # get each shape that should have been moved
        for shape_name in shape_names:
            s = shapes[0].find_shape_by_text(shape_name)  # type: VisioFile.Shape
            assert s  # check shape found
            assert s.x == expected_shape_locations[shape_name][0]
            assert s.y == expected_shape_locations[shape_name][1]


@pytest.mark.parametrize(("filename", "expected_connects"),
                         [('test4_connectors.vsdx', ["from 7 to 5", "from 7 to 2", "from 6 to 2", "from 6 to 1"]),
                          ])
def test_find_page_connects(filename: str, expected_connects: list):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: VisioFile.Page
        actual_connects = list()
        for c in page.connects:  # type: VisioFile.Connect
            actual_connects.append(f"from {c.from_id} to {c.to_id}")
        assert sorted(actual_connects) == sorted(expected_connects)


@pytest.mark.parametrize(("filename", "shape_id", "expected_shape_ids"),
                         [
                             ('test4_connectors.vsdx', "1", ["6"]),
                             ('test4_connectors.vsdx', "2", ["6", "7"]),
                          ])
def test_find_connected_shapes(filename: str, shape_id: str, expected_shape_ids: list):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: VisioFile.Page
        actual_connect_shape_ids = list()
        shape = page.find_shape_by_id(shape_id)
        for c_shape in shape.connected_shapes:  # type: VisioFile.Shape
            actual_connect_shape_ids.append(c_shape.ID)
        assert sorted(actual_connect_shape_ids) == sorted(expected_shape_ids)


@pytest.mark.parametrize(("filename", "shape_id", "expected_shape_ids", "expected_from", "expected_to", "expected_from_rels", "expected_to_rels"),
                         [
                             ('test4_connectors.vsdx', "1", ["6"], ["6"], ["1"], ['BeginX'], ['PinX']),
                             ('test4_connectors.vsdx', "2", ["6", "7"], ["6", "7"], ["2", "2"], ['BeginX', 'EndX'], ['PinX', 'PinX']),
                             ('test4_connectors.vsdx', "6", ["1", "2"], ["6", "6"], ["1", "2"], ['BeginX', 'EndX'], ['PinX', 'PinX']),
                             ('test4_connectors.vsdx', "7", ["5", "2"], ["7", "7"], ["5", "2"], ['BeginX', 'EndX'], ['PinX', 'PinX']),
                          ])
def test_find_connected_shape_relationships(filename: str, shape_id: str, expected_shape_ids: list, expected_from: list,
                                            expected_to: list, expected_from_rels: list, expected_to_rels: list):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: VisioFile.Page

        shape = page.find_shape_by_id(shape_id)
        shape_ids = [s.ID for s in shape.connected_shapes]
        from_ids = [c.from_id for c in shape.connects]
        to_ids = [c.to_id for c in shape.connects]
        from_rels = [c.from_rel for c in shape.connects]
        to_rels = [c.to_rel for c in shape.connects]

        assert sorted(shape_ids) == sorted(expected_shape_ids)
        assert sorted(from_ids) == sorted(expected_from)
        assert sorted(to_ids) == sorted(expected_to)
        assert sorted(from_rels) == sorted(expected_from_rels)
        assert sorted(to_rels) == sorted(expected_to_rels)


@pytest.mark.parametrize(("filename", "shape_a_id", "shape_b_id", "expected_connector_ids"),
                         [
                             ('test4_connectors.vsdx', "1", "2", ["6"]),
                             ('test4_connectors.vsdx', "1", "5", []),  # expect no connections
                          ])
def test_find_connectors_between_ids(filename: str, shape_a_id: str, shape_b_id: str,  expected_connector_ids: list):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: VisioFile.Page
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
        page = vis.pages[0]  # type: VisioFile.Page
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
        page = vis.pages[0]  # type: VisioFile.Page
        from_shape = page.find_shape_by_text(shape_a_text)
        to_shape = page.find_shape_by_text(shape_b_text)
        c = VisioFile.Connect.create(page=page, from_shape=from_shape, to_shape=to_shape)
        new_connector_id = c.ID
        conns_shown = []
        for conn in page.connects:
            if conn.connector_shape.ID not in conns_shown:
                print(f"conn between {[s.ID for s in conn.connector_shape.connected_shapes]} {conn.from_id}->{conn.to_id}:{VisioFile.pretty_print_element(conn.connector_shape.xml)}")
                conns_shown.append(conn.connector_shape.ID)
            print(VisioFile.pretty_print_element(conn.xml))
        vis.save_vsdx(out_file)

        # re-open saved file and check it is changed as expected
        with VisioFile(out_file) as vis:
            page = vis.pages[0]
            connector_ids = [c.connector_shape_id for c in page.connects]
            # new shape is referenced as connector_shape for new connector relationship
            assert new_connector_id in connector_ids
            # new shape exists in page
            c = page.find_shape_by_id(new_connector_id)
            print(f"conn_shape.line_to_x={c.line_to_x}")
            assert page.find_shape_by_id(new_connector_id)


@pytest.mark.parametrize(("filename", "shape_name"),
                         [("test1.vsdx", "Shape to copy"),
                          ("test4_connectors.vsdx", "Shape B")])
def test_vis_copy_shape(filename: str, shape_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_vis_copy_shape.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: VisioFile.Page
        # find and copy shape by name
        s = page.find_shape_by_text(shape_name)  # type: VisioFile.Shape
        assert s  # check shape found
        print(f"Found shape id:{s.ID}")
        max_id = page.max_id

        # note = this does add the shape, but prefer Shape.copy() as per next test which wraps this and returns Shape
        page.set_max_ids()
        new_shape = vis.copy_shape(shape=s.xml, page=page)

        assert new_shape  # check copy_shape returns xml

        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.attrib['ID']}")
        assert int(new_shape.attrib.get('ID')) > int(s.ID)
        assert int(new_shape.attrib.get('ID')) > max_id

        new_shape_id = new_shape.attrib['ID']
        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        s = page.find_shape_by_id(new_shape_id)
        assert s


@pytest.mark.parametrize(("filename", "shape_name"),
                         [("test1.vsdx", "Shape to copy"),
                          ("test4_connectors.vsdx", "Shape B")])
def test_shape_copy(filename: str, shape_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_shape_copy.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page =  vis.pages[0]  # type: VisioFile.Page
        # find and copy shape by name
        s = page.find_shape_by_text(shape_name)  # type: VisioFile.Shape
        assert s  # check shape found
        print(f"found {s.ID}")
        max_id = page.max_id

        new_shape = s.copy()
        assert new_shape  # check new shape exists

        print(f"original shape {type(s)} {s} {s.ID}")
        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.ID}")
        assert int(new_shape.ID) > int(s.ID)  # and new shape has > ID than original
        assert int(new_shape.ID) > max_id
        updated_text = s.text + " (new copy)"
        new_shape.text = updated_text  # update text of new shape

        new_shape_id = new_shape.ID
        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        s = page.find_shape_by_id(new_shape_id)
        assert s
        # check that new shape has expected text
        assert s.text == updated_text


@pytest.mark.parametrize(("filename", "shape_name"),
                         [("test1.vsdx", "Shape to copy"),
                          ("test2.vsdx", "Shape to copy")])
def test_copy_shape_other_page(filename: str, shape_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_copy_shape_other_page.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page =  vis.pages[0]  # type: VisioFile.Page
        page2 = vis.pages[1]  # type: VisioFile.Page
        page3 = vis.pages[2]  # type: VisioFile.Page
        # find and copy shape by name
        s = page.find_shape_by_text(shape_name)  # type: VisioFile.Shape
        assert s  # check shape found
        shape_text = s.text
        print(f"Found shape id:{s.ID}")

        new_shape = vis.copy_shape(shape=s.xml, page=page2)
        assert new_shape  # check copy_shape returns xml
        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.attrib['ID']}")
        page2_new_shape_id = new_shape.attrib['ID']

        new_shape = vis.copy_shape(shape=s.xml, page=page3)
        assert new_shape  # check copy_shape returns xml
        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.attrib['ID']}")
        page3_new_shape_id = new_shape.attrib['ID']

        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        page2 = vis.pages[1]
        s = page2.find_shape_by_id(page2_new_shape_id)
        assert s
        assert s.text == shape_text
        page3 = vis.pages[2]
        s = page3.find_shape_by_id(page3_new_shape_id)
        assert s
        assert s.text == shape_text


@pytest.mark.parametrize(("filename", "shape_name"),
                         [("test1.vsdx", "Shape to copy"),
                          ("test2.vsdx", "Shape to copy")])
def test_shape_copy_other_page(filename: str, shape_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_shape_copy_other_page.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page =  vis.pages[0]  # type: VisioFile.Page
        page2 = vis.pages[1]  # type: VisioFile.Page
        page3 = vis.pages[2]  # type: VisioFile.Page
        # find and copy shape by name
        s = page.find_shape_by_text(shape_name)  # type: VisioFile.Shape
        assert s  # check shape found
        shape_text = s.text
        print(f"Found shape id:{s.ID}")

        new_shape = s.copy(page2)
        assert new_shape  # check copy_shape returns xml
        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.ID}")
        page2_new_shape_id = new_shape.ID

        new_shape = s.copy(page3)
        assert new_shape  # check copy_shape returns xml
        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.ID}")
        page3_new_shape_id = new_shape.ID

        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        page2 = vis.pages[1]
        s = page2.find_shape_by_id(page2_new_shape_id)
        assert s
        assert s.text == shape_text
        page3 = vis.pages[2]
        s = page3.find_shape_by_id(page3_new_shape_id)
        assert s
        assert s.text == shape_text


@pytest.mark.parametrize(("filename", "expected_length"),
                         [('test5_master.vsdx', 1)])
def test_load_master_file(filename: str, expected_length: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        assert len(vis.master_pages) == expected_length


@pytest.mark.parametrize(("filename", "shape_text"),
                         [('test5_master.vsdx', "Shape B")])
def test_find_master_shape(filename: str, shape_text: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        master_page =  vis.master_pages[0]  # type: VisioFile.Page
        s = master_page.find_shape_by_text(shape_text)
        assert s


@pytest.mark.parametrize(("filename", "context"),
                         [("test_jinja.vsdx", {"date": datetime.now(), "scenario": "Scenario One", "x": 2, "y": 2}),
                          ("test_jinja.vsdx", {"date": datetime.now(), "scenario": "Scenario Two", "x": 2, "y": 2}),
                          ("test_jinja.vsdx", {"date": datetime.now(), "scenario": "Scenario Three", "x": 2, "y": 2}),
                          ])
def test_basic_jinja(filename: str, context: dict):

    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_basic_jinja.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]

        # each key string in context dict will be replaced with value, record against shape Ids for validation
        shape_id_values = dict()
        for k, v in context.items():
            shape_id = page.find_shape_by_text(k).ID
            shape_id_values[shape_id] = v
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and validate each shape id has expected text
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        for shape_id, text in shape_id_values.items():
            if type(text) is str:
                print(f"Testing that shape {shape_id} has text '{text}' in: {page.find_shape_by_id(shape_id).text}")
                assert str(text) in page.find_shape_by_id(shape_id).text


@pytest.mark.parametrize(("filename", "context", "shape_count"),
                         [("test_jinja.vsdx", {"date": datetime.now(), "scenario": "One", "x": 0, "y": 2}, 1),
                          ("test_jinja.vsdx", {"date": datetime.now(), "scenario": "Two", "x": 5, "y": 2}, 2),
                          ("test_jinja.vsdx", {"date": datetime.now(), "scenario": "Three", "x": 20, "y": 2}, 3),
                          ])
def test_jinja_if(filename: str, context: dict, shape_count: int):

    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_jinja_if_{context["scenario"]}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and validate each shape id has expected text
    with VisioFile(out_file) as vis:
        page = vis.pages[1]  # second page has the shapes with if statements
        count = len(page.sub_shapes())
        print(f"expected {shape_count} and found {count}")
        assert count == shape_count


@pytest.mark.parametrize(("filename", "context"),
                         [("test_jinja.vsdx", {"x": 2, "y": 2}),
                          ("test_jinja.vsdx", {"x": 3, "y": 4}),
                          ("test_jinja.vsdx", {"x": 12, "y": 9}),
                          ])
def test_jinja_calc(filename: str, context: dict):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_jinja_calc.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and validate each shape id has expected text
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        # check a shape exists with product of x and y values
        x_y = str(context['x'] * context['y'])
        assert page.find_shape_by_text(x_y)


@pytest.mark.parametrize(("filename", "context"),
                         [("test_jinja_loop.vsdx", {"date": datetime.now(), "scenario": "Scenario One", "test_list":[1,2,3]}),
                          ("test_jinja_loop.vsdx", {"date": datetime.now(), "scenario": "Scenario Two", "test_list":["One", "Two","Three"]}),
                          ("test_jinja_loop.vsdx", {"date": datetime.now(), "scenario": "Scenario Three", "test_list":[1,2,3,4,5,6]}),
                          ])
def test_basic_jinja_loop(filename: str, context: dict):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_basic_jinja.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]

        # each key string in context dict will be replaced with value, record against shape Ids for validation
        shape_id_values = dict()
        for k, v in context.items():
            shape_id = page.find_shape_by_text(k).ID
            shape_id_values[shape_id] = v
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and validate each shape id has expected text, and that a shape exists with each loop value
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        for shape_id, text in shape_id_values.items():
            if type(text) is str:
                print(f"Testing that shape {shape_id} has text '{text}' in: {page.find_shape_by_id(shape_id).text}")
                assert str(text) in page.find_shape_by_id(shape_id).text
            if type(text) is list:
                for item in text:
                    print(f"Testing that shape with text '{item}' exists")
                    assert page.find_shape_by_text(str(item))


@pytest.mark.parametrize(("filename", "context"),
                         [("test_jinja_inner_loop.vsdx", {"test_list":[[1, 2, 3], [1, 2, 3], [1, 2, 3]]}),
                          ("test_jinja_inner_loop.vsdx", {"test_list":["One", "Two","Three"]}),
                          ])
def test_jinja_inner_loop(filename: str, context: dict):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_jinja_inner_loop.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and validate each shape id has expected text, and that a shape exists with each loop value
    test_list = context["test_list"]
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        for o in test_list:
            for p in o:
                print(f"p={p}")
                s = page.find_shape_by_text(str(p))
                assert s


@pytest.mark.parametrize(("filename", "out_name", "context"),
                         [("test_jinja_loop_showif.vsdx", "1234", {"test_list": [1, 2, 3, 4]}),
                          ("test_jinja_loop_showif.vsdx", "3456", {"test_list": [3, 4, 5, 6]}),
                          ])
def test_jinja_loop_showif(filename: str, out_name: str,  context: dict):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}{out_name}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        for o in context['test_list']:
            s = page.find_shape_by_text(f"In this instance, o={o}")
            # test that {% show if o > 2 %} has worked
            print(f"for {o} found {s}")
            if o > 2:
                assert s  # check that shape is found (not removed by showif)
            else:
                assert not s  # check that shape is not found (has been removed by showif)


@pytest.mark.parametrize(
    ("filename", "context", "shape_id", "expected_x", "expected_text"),
    [("test_jinja_self_refs.vsdx", {"n": 1}, "1", 2.0, "This text should remain  and x should be 2.0\n"),
     ("test_jinja_self_refs.vsdx", {"n": 2}, "2", 4.0, "This shape sets x to n * 2\n"),
     ("test_jinja_self_refs.vsdx", {"n": 1}, "3", 1.0, "This shape sets x to 1 if n else 2\n"),
     ("test_jinja_self_refs.vsdx", {"n": 0}, "3", 2.0, "This shape sets x to 1 if n else 2\n"),
     ])
def test_jinja_self_refs(filename: str, context: dict, shape_id, expected_x, expected_text):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_jinja_self_refs.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: VisioFile.Page
        # there should be one shape on page 0
        shape = page.find_shape_by_id(shape_id)  # type: VisioFile.Shape
        print(f"DEBUG: ID={shape.ID} shape.text={shape.text}")
        print(f"DEBUG: ID={shape.ID} shape.x={shape.x}")
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and check shape has moved
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        shape = page.find_shape_by_id(shape_id)  # type: VisioFile.Shape
        print(f"DEBUG: ID={shape.ID} shape.text='{shape.text}' expected='{expected_text}'")
        print(f"DEBUG: ID={shape.ID} shape.x={shape.x}")
        assert shape.x == expected_x
        assert shape.text == expected_text


@pytest.mark.parametrize(
    ("filename", "context", "shape_id", "expected_y", "expected_text"),
    [("test_jinja_self_refs.vsdx", {"n": 2}, "4", 10.12368731806121, "This shape should move down by 1.0\n"),
     ("test_jinja_self_refs.vsdx", {"n": 1}, "5", 8.726049539918966, "This shape should move down by n\n"),
     ("test_jinja_self_refs.vsdx", {"n": 2}, "5", 7.726049539918966, "This shape should move down by n\n"),
     ])
def test_jinja_self_ref_calculations(filename: str, context: dict, shape_id, expected_y, expected_text):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_jinja_self_ref_calcs.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: VisioFile.Page
        # there should be one shape on page 0
        shape = page.find_shape_by_id(shape_id)  # type: VisioFile.Shape
        if shape:
            print(f"DEBUG: ID={shape.ID} shape.text={shape.text}")
            print(f"DEBUG: ID={shape.ID} shape.y={shape.y}")
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and check shape has moved
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        shape = page.find_shape_by_id(shape_id)  # type: VisioFile.Shape
        print(f"DEBUG: ID={shape.ID} shape.text='{shape.text}' expected='{expected_text}'")
        print(f"DEBUG: ID={shape.ID} shape.x={shape.y}")
        assert shape.y == expected_y
        assert shape.text == expected_text


@pytest.mark.parametrize(
    ("filename", "context", "expected_page_count", "expected_page_names"),
    [("test_jinja_page_showif.vsdx", {"show": True}, 2, ['Normal Page', 'Page2']),
     ("test_jinja_page_showif.vsdx", {"show": 1}, 2, ['Normal Page', 'Page2']),
     ("test_jinja_page_showif.vsdx", {"show": "true"}, 2, ['Normal Page', 'Page2']),
     ("test_jinja_page_showif.vsdx", {"show": 4.0}, 2, ['Normal Page', 'Page2']),
     ("test_jinja_page_showif.vsdx", {"show": False}, 2, ['Normal Page', 'Page3']),
     ("test_jinja_page_showif.vsdx", {"show": []}, 2, ['Normal Page', 'Page3']),
     ("test_jinja_page_showif.vsdx", {"show": {}}, 2, ['Normal Page', 'Page3']),
     ])
def test_jinja_page_showif(filename: str, context: dict, expected_page_count, expected_page_names):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_show_{context["show"]}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        print(f"len(vis.pages)={len(vis.pages)} context={context}")
        for p in vis.pages: # type: VisioFile.Page
            print(f"page:{p.name}")
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and check shape has moved
    with VisioFile(out_file) as vis:
        page_names = []
        for p in vis.pages: # type: VisioFile.Page
            print(f"page:{p.name}")
            page_names.append(p.name)
        assert len(vis.pages) == expected_page_count
        assert page_names == expected_page_names


@pytest.mark.parametrize(("filename", "shape_elements"), [("test1.vsdx", 4), ("test2.vsdx", 14), ("test3_house.vsdx", 10)])
def test_xml_findall_shapes(filename: str, shape_elements: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        # find all Shape elements
        xml = page.xml.getroot()
        xpath = f".//{namespace}Shape"
        elements = xml.findall(xpath)
        print(f"{xpath} returns {len(elements)} elements vs {shape_elements}")
        assert len(elements) == shape_elements


@pytest.mark.parametrize(("filename", "group_shape_elements"), [("test1.vsdx", 0), ("test2.vsdx", 3), ("test3_house.vsdx", 2)])
def test_xml_findall_group_shapes(filename: str, group_shape_elements: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        # find all Shape elements where attribute Type='Group'
        xml = page.xml.getroot()
        xpath = f".//{namespace}Shape[@Type='Group']"
        elements = xml.findall(xpath)
        print(f"{xpath} returns {len(elements)} elements vs {group_shape_elements}")
        assert len(elements) == group_shape_elements


@pytest.mark.parametrize(("filename", "page_index"), [("test1.vsdx", 0)])
def test_remove_page_by_index(filename: str, page_index: int):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_remove_page.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page_count = len(vis.pages)
        vis.remove_page_by_index(page_index)
        vis.save_vsdx(out_file)

    # re-open file and confirm it has one less page
    with VisioFile(out_file) as vis:
        assert len(vis.pages) == page_count - 1


@pytest.mark.parametrize(("filename"),
                         [('test5_master.vsdx')])
def test_master_inheritance(filename: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        master_page = vis.master_pages[0]  # type: VisioFile.Page
        shape_a = page.find_shape_by_text('Shape A')  # type: VisioFile.Shape
        shape_b = page.find_shape_by_text('Shape B')  # type: VisioFile.Shape

        assert shape_a
        assert shape_b

        # test inheritance of cell values
        assert shape_a.width == 1
        assert shape_a.height == 0.75

        # test inheritance for subshapes
        sub_shape_a = shape_a.sub_shapes()[0]
        assert sub_shape_a.cell_value('LineWeight') == '0.01875'


@pytest.mark.parametrize(("filename"),
                         [('test5_master.vsdx')])
def test_master_inheritance_setters(filename: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_master_inheritance_setters.vsdx')
    # f'{basedir}out{os.sep}{filename[:-5]}_test_master_inheritance_setters.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)

        shape_a = page.find_shape_by_text('Shape A')
        shape_a.line_weight = 0.5
        shape_a.line_color = '#00FF00'

        sub_shape_b = page.find_shape_by_text('Shape B').sub_shapes()[0]
        sub_shape_b.line_weight = 0.5
        sub_shape_b.line_color = '#00FF00'

        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page(0)

        shape_a = page.find_shape_by_text('Shape A')
        assert shape_a.line_weight == 0.5
        assert shape_a.line_color == '#00FF00'

        sub_shape_b = page.find_shape_by_text('Shape B').sub_shapes()[0]
        assert sub_shape_b.line_weight == 0.5
        assert sub_shape_b.line_color == '#00FF00'


@pytest.mark.parametrize(("filename", "shape_text"),
                         [("test_master.vsdx", "Page Shape"),
                          ("test_master.vsdx", "Master Shape A"),
                          ("test_master.vsdx", "Master Shape B"),
                          ("test_master.vsdx", "Master B with updated text"), ])
def test_master_find_shapes(filename: str, shape_text: str):
    # Check that shape with text can be found - whether in page, master or overridden in master
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        shape = page.find_shape_by_text(shape_text)
        assert shape  # shape found


@pytest.mark.parametrize(("filename", "shape_text", "has_master", "inherits_text"),
                         [("test_master.vsdx", "Page Shape", False, False),
                          ("test_master.vsdx", "Master Shape A", True, True),
                          ("test_master.vsdx", "Master Shape B", True, True),
                          ("test_master.vsdx", "Master B with updated text", True, False), ])
def test_master_check_text_inheritance(filename: str, shape_text: str, has_master: bool, inherits_text: bool):
    # Check that shape with text can be found and that it has a master (or not) and inherits text (or not)
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: VisioFile.Page
        shape = page.find_shape_by_text(shape_text)

        if shape.master_shape:
            assert has_master
            assert (shape.text == shape.master_shape.text) == inherits_text
        else:
            assert not has_master


@pytest.mark.parametrize(("filename", "master_shape_text", "number_shapes"),
                         [('test_master.vsdx', 'Master Shape A', 1),
                          ('test_master.vsdx', 'Master Shape B', 2),
                          ])
def test_find_shapes_by_master_id(filename: str, master_shape_text: str, number_shapes: int):
    # test that a shapes master has 'number_shapes' shapes that inherit from it
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)
        shape = page.find_shape_by_text(master_shape_text)  # get shape which has a master
        # print(f"master_shape: {master_shape} {master_shape_.master_shape_ID}")
        shapes = page.find_shapes_with_same_master(shape)
        print(f"shapes: {shapes}")
        assert len(shapes) == number_shapes


@pytest.mark.parametrize(("filename", "master_shape_search_text", "shape_search_text", "expect_inherit"),
                         [('test_master.vsdx', 'Master Shape A', 'Master Shape A', True),
                          ('test_master.vsdx', 'Master Shape A', 'Master B with updated text', False),
                          ('test_master.vsdx', 'Master B with updated text', 'Master B with updated text', False),
                          ('test_master.vsdx', 'Master Shape B', 'Master Shape B', True),
                          ('test_master.vsdx', 'Master Shape A', 'Page Shape', False),
                          ])
def test_master_inheritance_master_shape_set_text(filename: str, master_shape_search_text: str, shape_search_text: str,
                                                  expect_inherit: bool):
    # test that when a master shape text is updated, only shapes that inherit that master shapes text are affected
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_master_inheritance_master_shape_set_text.vsdx')
    master_text = "Updated Master Shape Text"
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)

        # get master shape of Shape with search text
        master_shape = page.find_shapes_by_text(master_shape_search_text)[0].master_shape
        shape = page.find_shape_by_text(shape_search_text)
        shape_id = shape.ID
        shape_text = shape.text
        master_shape.text = master_text

        if expect_inherit:
            assert shape.text == master_text  # inherited from master change
        else:
            assert shape.text == shape_text  # unchanged

        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page(0)

        # get same shape and confirm inheritance works as expected
        shape = page.find_shape_by_id(shape_id)
        if expect_inherit:
            assert shape.text == master_text  # inherited from master change
        else:
            assert shape.text == shape_text  # unchanged


@pytest.mark.parametrize(('filename'),
                         [('test1.vsdx')])
def test_add_page(filename: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_add_page.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        number_pages = len(vis.pages)

        new_page = vis.add_page()
        assert new_page
        assert len(vis.pages) == number_pages + 1

        new_page_name = new_page.name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(new_page_name)
        assert page


@pytest.mark.parametrize(('filename', 'page_name'),
                         [
                            ('test1.vsdx', 'newname'),
                            ('test1.vsdx', 'Page-1'),
                         ])
def test_add_page_name(filename: str, page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_add_page_name_{page_name}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        number_pages = len(vis.pages)

        new_page = vis.add_page(page_name)
        assert new_page
        assert len(vis.pages) == number_pages + 1

        new_page_name = new_page.name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(new_page_name)
        assert page


@pytest.mark.parametrize(('filename', 'index', 'page_name'),
                         [
                            ('test1.vsdx', 1, None),
                            ('test1.vsdx', 1, 'newname'),
                            ('test1.vsdx', 1, 'Page-1'),
                         ])
def test_add_page_at(filename: str, index: int, page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_add_page_at_{page_name}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        number_pages = len(vis.pages)

        new_page = vis.add_page_at(index, page_name)
        assert new_page
        assert len(vis.pages) == number_pages + 1

        new_page_name = new_page.name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(new_page_name)
        assert page


@pytest.mark.parametrize(('filename', 'index', 'page_name'),
                         [
                            ('test1.vsdx', -1, None),
                            ('test1.vsdx', 0, 'newname'),
                            ('test1.vsdx', 2, 'Page-1'),
                            ('test2.vsdx', 2, 'Page-1'),
                            ('test4_connectors.vsdx', 0, 'Page-1'),
                         ])
def test_copy_page(filename: str, index: int, page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_copy_page_{page_name}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        number_pages = len(vis.pages)

        page = vis.pages[0]  # type: VisioFile.Page
        new_page = vis.copy_page(page, index=index, name=page_name)
        assert new_page
        assert len(vis.pages) == number_pages + 1

        new_page_name = new_page.name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(new_page_name)
        assert page


@pytest.mark.parametrize(('filename', 'page_index_to_copy', 'in_page_name', 'out_page_name'),
                         [
                             ('test1.vsdx', 0, None, 'Page-1-1'),
                             ('test1.vsdx', 0,  'newname', 'newname'),
                             ('test1.vsdx', 0, 'Page-1', 'Page-1-1'),
                             ('test1.vsdx', 1, None, 'Page-2-1'),
                             ('test1.vsdx', 1,  'newname', 'newname'),
                             ('test1.vsdx', 1, 'Page-1', 'Page-1-1'),
                         ])
def test_copy_page_naming(filename: str, page_index_to_copy:int, in_page_name: str, out_page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_copy_page_naming.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index_to_copy]  # type: VisioFile.Page
        new_page = vis.copy_page(page, name=in_page_name)
        print(f"in_page_name:{in_page_name} out_page_name:{out_page_name} actual:{new_page.name}")
        assert new_page.name == out_page_name  # check new page has expected name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(out_page_name)
        assert page  # check that page name persists through file save and open


@pytest.mark.parametrize(('filename', 'page_index_to_copy', 'page_position', 'out_page_index'),
                         [
                             ('test1.vsdx', 0, PagePosition.LAST, 3),
                             ('test1.vsdx', 0, PagePosition.BEFORE, 0),
                             ('test1.vsdx', 0, PagePosition.AFTER, 1),
                             ('test1.vsdx', 1, PagePosition.LAST, 3),
                             ('test1.vsdx', 1, PagePosition.BEFORE, 1),
                             ('test1.vsdx', 1, PagePosition.AFTER, 2),
                         ])
def test_copy_page_positions(filename: str, page_index_to_copy:int, page_position: PagePosition, out_page_index: int):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_copy_page_position.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index_to_copy]  # type: VisioFile.Page
        new_page = vis.copy_page(page, index=page_position)
        index = vis.pages.index(new_page)
        print(f"page_index_to_copy:{page_index_to_copy} page_position:{page_position} out_page_index:{out_page_index} actual:{index}")
        assert index == out_page_index  # check new page has expected index
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(new_page.name)
        index = vis.pages.index(page)
        assert index == out_page_index  # check that page location persists through file save and open


@pytest.mark.parametrize(("filename"),
                         [("test1.vsdx"),
                          ])
def test_app_xml_page_names(filename: str):
    # test that page names in app.xml matches page names loaded
    with VisioFile(os.path.join(basedir, filename)) as vis:
        HeadingPairs = vis.app_xml.getroot().find(f'{ext_prop_namespace}HeadingPairs')
        i4 = HeadingPairs.find(f'.//{vt_namespace}i4')
        num_pages = int(i4.text)
        assert num_pages == len(vis.pages)

        # check page names from pages is same as page names from app.xml
        page_names = [p.name for p in vis.pages]
        TitlesOfParts = vis.app_xml.getroot().find(f'{ext_prop_namespace}TitlesOfParts')
        vector = TitlesOfParts.find(f'{vt_namespace}vector')
        app_xml_page_names = []
        for lpstr in vector.findall(f'{vt_namespace}lpstr'):
            page_name = lpstr.text
            app_xml_page_names.append(page_name)
        assert page_names == app_xml_page_names


@pytest.mark.parametrize(("filename", "new_page_name", "location"),
                         [("test1.vsdx", "new_page", 0),
                          ("test1.vsdx", "new_page", 1),
                          ("test2.vsdx", "new_page", 0),
                          ])
def test_app_xml_page_names_after_add_page(filename: str, new_page_name: str, location: int):
    # test that page names in app.xml matches page names loaded
    with VisioFile(os.path.join(basedir, filename)) as vis:
        if location is not None:
            vis.add_page(new_page_name)
        else:
            vis.add_page_at(location, new_page_name)

        HeadingPairs = vis.app_xml.getroot().find(f'{ext_prop_namespace}HeadingPairs')
        i4 = HeadingPairs.find(f'.//{vt_namespace}i4')
        num_pages = int(i4.text)
        assert num_pages == len(vis.pages)

        # check page names from pages is same as page names from app.xml
        page_names = [p.name for p in vis.pages]
        TitlesOfParts = vis.app_xml.getroot().find(f'{ext_prop_namespace}TitlesOfParts')
        vector = TitlesOfParts.find(f'{vt_namespace}vector')
        app_xml_page_names = []
        for lpstr in vector.findall(f'{vt_namespace}lpstr'):
            page_name = lpstr.text
            app_xml_page_names.append(page_name)
        assert page_names == app_xml_page_names


@pytest.mark.parametrize(("filename", "remove_index"),
                         [("test1.vsdx", 0),
                          ("test1.vsdx", 1),
                          ("test2.vsdx", 0),
                          ])
def test_app_xml_page_names_after_remove_page(filename: str, remove_index: int):
    # test that page names in app.xml matches page names loaded
    with VisioFile(os.path.join(basedir, filename)) as vis:
        vis.remove_page_by_index(remove_index)

        HeadingPairs = vis.app_xml.getroot().find(f'{ext_prop_namespace}HeadingPairs')
        i4 = HeadingPairs.find(f'.//{vt_namespace}i4')
        num_pages = int(i4.text)
        assert num_pages == len(vis.pages)

        # check page names from pages is same as page names from app.xml
        page_names = [p.name for p in vis.pages]
        TitlesOfParts = vis.app_xml.getroot().find(f'{ext_prop_namespace}TitlesOfParts')
        vector = TitlesOfParts.find(f'{vt_namespace}vector')
        app_xml_page_names = []
        for lpstr in vector.findall(f'{vt_namespace}lpstr'):
            page_name = lpstr.text
            app_xml_page_names.append(page_name)
        assert page_names == app_xml_page_names


@pytest.mark.parametrize("filename, page_index, shape_text, expected_coords", [
    ("test7_with_connector.vsdx", 0, "Tall Box", [('RelMoveTo', [('X', '0', None), ('Y', '0', None)]), ('RelLineTo', [('X', '1', None), ('Y', '0', None)]), ('RelLineTo', [('X', '1', None), ('Y', '1', None)]), ('RelLineTo', [('X', '0', None), ('Y', '1', None)]), ('RelLineTo', [('X', '0', None), ('Y', '0', None)])]),
    ("test1.vsdx", 0, "Shape Text", [('RelMoveTo', [('X', '0', None), ('Y', '0', None)]), ('RelLineTo', [('X', '1', None), ('Y', '0', None)]), ('RelLineTo', [('X', '1', None), ('Y', '1', None)]), ('RelLineTo', [('X', '0', None), ('Y', '1', None)]), ('RelLineTo', [('X', '0', None), ('Y', '0', None)])]),
    ("test2.vsdx", 2, "Already here", [('Ellipse', [('X', '0.2460629842730571', 'Width*0.5'), ('Y', '0.2519684958956088', 'Height*0.5'), ('A', '0.4921259685461141', 'Width*1'), ('B', '0.2519684958956088', 'Height*0.5'), ('C', '0.2460629842730571', 'Width*0.5'), ('D', '0.5039369917912175', 'Height*1')])]),
    ("test4_connectors.vsdx", 1, "A to B", [('MoveTo', [('X', '0', None), ('Y', '0.09842519685039441', None)]), ('LineTo', [('X', '0.6358267353988465', None), ('Y', '0.09842519685039441', None)])]),
])
def test_get_shape_geometry(filename: str, page_index: str, shape_text: str, expected_coords: list):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index]
        shape = page.find_shape_by_text(shape_text)

        coords = [(r.row_type, [(c.name, c.value, c.func) for c in r.cells.values()]) for r in shape.geometry.rows.values()]
        print(f"coords={coords}")
        print(VisioFile.pretty_print_element(shape.xml))
        if shape.master_shape:
            print(VisioFile.pretty_print_element(shape.master_shape.xml))
        assert coords == expected_coords