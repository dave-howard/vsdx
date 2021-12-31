import pytest
from vsdx import VisioFile, namespace, vt_namespace, ext_prop_namespace, PagePosition, Media
from datetime import datetime
import os
from typing import List

# code to get basedir of this test file in either linux/windows
basedir = os.path.dirname(os.path.relpath(__file__))


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


# DataProperty

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


# Connectors

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
