"""Tests for Shape class"""
import os
import pytest

import vsdx
from vsdx import DataProperty
from vsdx import Page  # for typing
from vsdx import Shape
from vsdx import VisioFile

# code to get basedir of this test file in either linux/windows
basedir = os.path.dirname(os.path.relpath(__file__))


@pytest.mark.parametrize("filename, shape_id, expected_text", [
    ("test3_house.vsdx", "1", "Shape Text\n"),
    ("test3_house.vsdx", "11", "Shape to remove\n"),
    ("test2.vsdx", "9", "Group shape text\n"),
    ("test10_nested_shapes.vsdx", "5", "Shape 1.2.1\n"),
    ("test_master_multiple_child_shapes.vsdx", "3", "AWS Step Functions workflow \n")
    ])
def test_get_shape_text(filename: str, shape_id: str, expected_text: str):
    # Check that a specific shape on a page has expected text value
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        shape = page.find_shape_by_id(shape_id)
        # check that shape has expected text
        assert shape.text == expected_text


@pytest.mark.parametrize("filename, shape_id, expected_text", [
    ("test3_house.vsdx", "1", "Shape Text\n"),
    ("test3_house.vsdx", "11", "Shape to remove\n"),
    ("test2.vsdx", "9", "Group shape text\n"),
    ("test10_nested_shapes.vsdx", "5", "Shape 1.2.1\n"),
    ])
def test_set_shape_text(filename: str, shape_id: str, expected_text: str):
    # Check that a specific shape on a page has expected text value after it is updated
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        shape = page.find_shape_by_id(shape_id)
        # check that shape has expected text
        assert shape.text == expected_text
        shape.text = expected_text + 'changed'
        shape = page.find_shape_by_id(shape_id)
        assert shape.text == expected_text + 'changed'


@pytest.mark.parametrize("filename, attr, attr_value, expected_id", [
    ("test3_house.vsdx", "NameU", "House", "7"),
    ])
def test_get_shape_attr_value(filename: str, attr: str, attr_value: str, expected_id: str):
    # Check that a specific shape on a page has expected text value
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        shape = page.find_shape_by_attr(attr, attr_value)
        # check that shape has expected text
        assert shape.ID == expected_id


@pytest.mark.parametrize("filename, shape_id, child_count", [
    ("test3_house.vsdx", "7", 3),
    ("test3_house.vsdx", "11", 3),
    ("test2.vsdx", "9", 3),
    ("test10_nested_shapes.vsdx", "7", 2),
    ])
def test_get_shape_child_shapes(filename: str, shape_id: str, child_count: int):
    # Check that page has expected number of top level shapes
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        shape = page.find_shape_by_id(shape_id)

        # check that page has expected number of child shapes
        assert len(shape.child_shapes) == child_count


@pytest.mark.parametrize("filename, shape_id, all_count", [
    ("test3_house.vsdx", "7", 3),
    ("test3_house.vsdx", "11", 3),
    ("test2.vsdx", "9", 3),
    ("test10_nested_shapes.vsdx", "7", 6),
    ])
def test_get_shape_all_shapes(filename: str, shape_id: str, all_count: int):
    # Check that page has expected number of top level shapes
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page

        shape = page.find_shape_by_id(shape_id)
        for shape in shape.all_shapes:
            print(shape)
        # check that page has expected number of child shapes
        assert len(shape.all_shapes) == all_count


@pytest.mark.parametrize(
    "filename, expected_locations",
    [("test1.vsdx", "1.33,10.66 4.13,10.66 6.94,10.66 2.33,9.02 "),
     ("test2.vsdx", "2.33,8.72 1.33,10.66 4.13,10.66 5.91,8.72 1.61,8.58 3.25,8.65 ")])
def test_shape_locations(filename: str, expected_locations: str):
    print("=== list_shape_locations ===")
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        shapes = page.child_shapes
        locations = ""
        for shape in shapes:  # type: Shape
            locations += f"{shape.x:.2f},{shape.y:.2f} "
        print(f"Expected:{expected_locations}")
        print(f"  Actual:{locations}")
    assert locations == expected_locations


@pytest.mark.parametrize(
    "filename, shape_id, expected_center",
    [
        ("test1.vsdx", "1", (1.332677148526936, 10.65551182326173)),
        ("test2.vsdx", "2", (1.082677148526936, 0.7874015625650443)),  # center of a group shape
        ("test2.vsdx", "16", (1.6903102768832179, 8.188976116607332)),  # test center of a line
    ])
def test_shape_center(filename: str, shape_id: str, expected_center: str):

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        shape = page.find_shape_by_id(shape_id)

        assert shape.center_x_y == expected_center


@pytest.mark.parametrize("filename", ["test2.vsdx", "test3_house.vsdx"])
def test_remove_shape(filename: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_shape_removed.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        # get shape to remove
        shape = vis.pages[0].find_shape_by_text('Shape to remove')  # type: Shape
        assert shape  # check shape found
        shape.remove()
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        # get shape that should have been removed
        shape = vis.pages[0].find_shape_by_text('Shape to remove')  # type: Shape
        assert shape is None  # check shape not found


@pytest.mark.parametrize(("filename", "shape_names", "shape_locations"),
                         [("test1.vsdx", {"Shape to remove"}, {(1.0, 1.0)})])
def test_set_shape_location(filename: str, shape_names: set, shape_locations: set):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_set_shape_location.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        # move shapes in list
        for (shape_name, x_y) in zip(shape_names, shape_locations):
            shape = vis.pages[0].find_shape_by_text(shape_name)  # type: Shape
            assert shape  # check shape found
            assert shape.x
            assert shape.y
            print(f"Moving shape '{shape_name}' from {shape.x}, {shape.y} to {x_y}")
            shape.x = x_y[0]
            shape.y = x_y[1]
        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        # get each shape that should have been moved
        for (shape_name, x_y) in zip(shape_names, shape_locations):
            shape = vis.pages[0].find_shape_by_text(shape_name)  # type: Shape
            assert shape  # check shape found
            assert shape.x == x_y[0]
            assert shape.y == x_y[1]


@pytest.mark.parametrize(("filename", "page_index", "shape_names", "shape_x_y_deltas"),
                         [("test1.vsdx", 0, {"Shape to remove"}, {(1.0, 1.0)}),
                          ("test4_connectors.vsdx", 0, {"Shape B"}, {(1.0, 1.0)}),
                          ("test4_connectors.vsdx", 1, {"B to C"}, {(1.0, 1.0)})])
def test_move_shape(filename: str, page_index: int, shape_names: set, shape_x_y_deltas: set):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_move_shape.vsdx')
    expected_shape_locations = {}

    with VisioFile(os.path.join(basedir, filename)) as vis:
        # move shapes in list
        for (shape_name, x_y) in zip(shape_names, shape_x_y_deltas):
            shape = vis.pages[page_index].find_shape_by_text(shape_name)  # type: Shape
            assert shape  # check shape found
            assert shape.x
            assert shape.y
            expected_shape_locations[shape_name] = (shape.x + x_y[0], shape.y + x_y[1])
            shape.move(x_y[0], x_y[1])

        vis.save_vsdx(out_file)
    print(f"expected_shape_locations={expected_shape_locations}")

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        # get each shape that should have been moved
        for shape_name in shape_names:
            shape = vis.pages[page_index].find_shape_by_text(shape_name)  # type: Shape
            assert shape  # check shape found
            assert shape.x == expected_shape_locations[shape_name][0]
            assert shape.y == expected_shape_locations[shape_name][1]


@pytest.mark.parametrize(("filename", "shape_name"),
                         [("test1.vsdx", "Shape to copy"),
                          ("test4_connectors.vsdx", "Shape B")])
def test_shape_copy(filename: str, shape_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_shape_copy.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page =  vis.pages[0]  # type: Page
        # find and copy shape by name
        shape = page.find_shape_by_text(shape_name)  # type: Shape
        assert shape  # check shape found
        print(f"found {shape.ID}")
        max_id = page.max_id

        new_shape = shape.copy()
        assert new_shape  # check new shape exists

        print(f"original shape {type(shape)} {shape} {shape.ID}")
        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.ID}")
        assert int(new_shape.ID) > int(shape.ID)  # and new shape has > ID than original
        assert int(new_shape.ID) > max_id
        updated_text = shape.text + " (new copy)"
        new_shape.text = updated_text  # update text of new shape
        assert page.find_shape_by_text(updated_text)
        new_shape_id = new_shape.ID
        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        shape = page.find_shape_by_id(new_shape_id)
        assert shape
        # check that new shape has expected text
        assert shape.text == updated_text


@pytest.mark.parametrize(("filename", "shape_name"),
                         [("test1.vsdx", "Shape to copy"),
                          ("test2.vsdx", "Shape to copy")])
def test_shape_copy_other_page(filename: str, shape_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_shape_copy_other_page.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page =  vis.pages[0]  # type: Page
        page2 = vis.pages[1]  # type: Page
        page3 = vis.pages[2]  # type: Page
        # find and copy shape by name
        shape = page.find_shape_by_text(shape_name)  # type: Shape
        assert shape  # check shape found
        shape_text = shape.text
        print(f"Found shape id:{shape.ID}")

        new_shape = shape.copy(page2)
        assert new_shape  # check copy_shape returns xml
        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.ID}")
        page2_new_shape_id = new_shape.ID

        new_shape = shape.copy(page3)
        assert new_shape  # check copy_shape returns xml
        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.ID}")
        page3_new_shape_id = new_shape.ID

        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        page2 = vis.pages[1]
        shape = page2.find_shape_by_id(page2_new_shape_id)
        assert shape
        assert shape.text == shape_text
        page3 = vis.pages[2]
        shape = page3.find_shape_by_id(page3_new_shape_id)
        assert shape
        assert shape.text == shape_text


# DataProperty

@pytest.mark.parametrize(("filename", "page_index", "shape_name"),
                         [("test1.vsdx", 0, "Shape Text"),
                          ("test6_shape_properties.vsdx", 2, "A")
                          ])
def test_get_shape_data_properties_type_is_dict_of_data_property(filename: str, page_index: int,
                                                                 shape_name: str):
    """ check that Shape.data_properties is a dict of ShapeProperty """
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shape = vis.pages[page_index].find_shape_by_text(shape_name)  # type: Shape

        props = shape.data_properties

        # check overall type
        assert isinstance(props, dict)

        # check each item type
        for prop in props.values():
            assert isinstance(prop, DataProperty)


@pytest.mark.parametrize(("filename", "page_index", "shape_name", "property_dict"),
                         [("test1.vsdx", 0, "Shape Text", {"my_property_label": "property value",
                                                           "my_second_property_label": "another value",
                                                           "Network Name": "Box01"}),
                          ("test6_shape_properties.vsdx", 2, "A", {"master_Prop": "master prop value"}),
                          ("test6_shape_properties.vsdx", 2, "B", {"master_Prop": "master prop value",
                                                                   "shape_prop": "shape property value"}),
                          ("test6_shape_properties.vsdx", 2, "C", {"master_Prop": "override"}),
                          ("test6_shape_properties.vsdx", 2, "D", {"LongProp": 'value not in an "attrib"'}),
                          ])
def test_get_shape_data_properties(filename: str, page_index: int, shape_name: str, property_dict: dict):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shape = vis.pages[page_index].find_shape_by_text(shape_name)  # type: Shape

        props = shape.data_properties
        print(props)
        # check lengths are same
        assert len(props) == len(property_dict)

        # check each key/value is same as expected
        for property_label in props.keys():
            prop = props[property_label]
            print(f"prop: lbl:'{prop.label}' name:'{prop.name}': val:'{prop.value}'")
            assert prop.value == property_dict.get(property_label)


@pytest.mark.parametrize(("filename", "page_index", "shape_name", "property_dict"),
                         [#("test1.vsdx", 0, "Shape Text", {"my_property_label": "1",
                          #                                 "my_second_property_label": "2",
                          #                                 "Network Name": "3"}),
                          #("test6_shape_properties.vsdx", 2, "A", {"master_Prop": "1"}),
                          #("test6_shape_properties.vsdx", 2, "B", {"master_Prop": "1",
                          #                                         "shape_prop": "2"}),
                          #("test6_shape_properties.vsdx", 2, "C", {"master_Prop": "1"}),
                          #("test6_shape_properties.vsdx", 2, "D", {"LongProp": '1"'}),
                          ("test_shape_with_field.vsdx", 0, "Here is field", {"field_label": 'updated field value'}),
                          ])
def test_set_shape_data_properties(filename: str, page_index: int, shape_name: str, property_dict: dict):
    """Check that we can set a prop value, and this change persists through save and load"""
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_set_shape_data_properties.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shape = vis.pages[page_index].find_shape_by_text(shape_name)

        props = shape.data_properties

        # check each key/value is not already set to required value
        for property_label in property_dict.keys():
            prop = props[property_label]
            print(f"prop: lbl:'{prop.label}' name:'{prop.name}': val:'{prop.value}'")
            assert prop.value != property_dict.get(property_label)
            print(f"shape.text={shape.text}")
            print(vis.pretty_print_element(shape.xml))

        # check each key/value is expected after being set
        for property_label in property_dict.keys():
            prop = props[property_label]
            prop.value = property_dict.get(property_label)
            print(f"checking prop after set: lbl:'{prop.label}' name:'{prop.name}': val:'{prop.value}'")
            assert prop.value == property_dict.get(property_label)

        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        shape = vis.pages[page_index].find_shape_by_text(shape_name)  # type: Shape
        # check each key/value is not already set to required value
        for property_label in property_dict.keys():
            prop = props[property_label]
            print(f"checking prop after load: lbl:'{prop.label}' name:'{prop.name}': val:'{prop.value}'")
            assert prop.value == property_dict.get(property_label)


@pytest.mark.parametrize(("filename", "page_index", "container_shape_name", "expected_shape_name", "property_label"),
                         [("test6_shape_properties.vsdx", 1, "Container", "Shape A", "label_one"),
                          ])
def test_find_sub_shape_by_data_property_label(filename: str, page_index: int, container_shape_name: str, expected_shape_name: list, property_label: str):
    """Test that we can find a shape and it's text inside a container shape based on the sub shapes labels"""
    with VisioFile(os.path.join(basedir, filename)) as vis:
        container_shape = vis.pages[page_index].find_shape_by_text(container_shape_name)

        shape = container_shape.find_shape_by_property_label(property_label)

        # test that a shape is returned
        assert shape

        # test that the shape returned has the expected names
        assert shape.text.replace('\n', '') == expected_shape_name


@pytest.mark.parametrize(("filename", "page_index", "property_label", "expected_shape_text"),
                          [("test_master_multiple_child_shapes.vsdx", 0, "title",
                            "AWS Step Functions workflow ")
                          ])
def test_find_shape_by_data_property_label(filename: str, page_index: int, property_label: str,
                                           expected_shape_text: str):
    """Test that we can find a shape and it's text inside a page based on the shapes labels"""
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shape = vis.pages[page_index].find_shape_by_property_label(property_label)

        # test that a shape is returned
        assert shape

        # test that shape has a DataProperty with expected property label
        assert [prop for prop in shape.data_properties.values() if prop.label == property_label]

        # test that the shape returned has the expected text
        assert shape.text.replace('\n', '') == expected_shape_text


# Connectors

@pytest.mark.parametrize(("filename", "shape_id", "expected_shape_ids"),
                         [
                             ('test4_connectors.vsdx', "1", ["6"]),
                             ('test4_connectors.vsdx', "2", ["6", "7"]),
                          ])
def test_find_connected_shapes(filename: str, shape_id: str, expected_shape_ids: list):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: Page
        actual_connect_shape_ids = list()
        shape = page.find_shape_by_id(shape_id)
        for c_shape in shape.connected_shapes:  # type: Shape
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
        page = vis.pages[0]  # type: Page

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


@pytest.mark.parametrize("filename_1, page_index_1, shape_text_1, filename_2, page_index_2, shape_text_2, are_equal", [
    ("test1.vsdx", 0, "Shape Text", "test1.vsdx", 0, "Shape Text", True),
    ("test1.vsdx", 0, "Shape Text", "test2.vsdx", 0, "Shape Text", False),
    ])
def test_shape_equality(filename_1, page_index_1, shape_text_1, filename_2, page_index_2, shape_text_2, are_equal):
    with VisioFile(os.path.join(basedir, filename_1)) as vis:
        page = vis.pages[page_index_1]
        shape_1 = page.find_shape_by_text(shape_text_1)

    with VisioFile(os.path.join(basedir, filename_2)) as vis:
        page = vis.pages[page_index_2]
        shape_2 = page.find_shape_by_text(shape_text_2)

    assert (shape_1 == shape_2) == are_equal


@pytest.mark.parametrize("filename, page_index, shape_text, expected_bounds", [
    ("test1.vsdx", 0, "Shape Text", ['0.25', '9.87', '2.42', '11.44']),  # standard shape
    ("test2.vsdx", 0, "Sub-shape 2", ['0.54', '0.17', '1.62', '0.51']),  # sub shape in a group
    ("test2.vsdx", 0, "Scenario:", ['0.73', '8.19', '2.49', '8.97']),  # line
    ])
def test_shape_bounds(filename, page_index, shape_text, expected_bounds):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index]
        shape = page.find_shape_by_text(shape_text)
        print([f"{n:.2f}" for n in shape.bounds])
        assert [f"{n:.2f}" for n in shape.bounds] == expected_bounds


@pytest.mark.parametrize("filename, page_index", [
    ("test1.vsdx", 0),
    ("test1.vsdx", 2),
    ("test2.vsdx", 0),
    ])
def test_all_shape_bounds(filename, page_index):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index]
        for s in page.all_shapes:
            # checl all shapes have bounds and relative bounds
            print(s.ID, s.bounds, s.relative_bounds)
            assert s.bounds
            assert s.relative_bounds


@pytest.mark.parametrize("filename, page_index, shape_text, expected_bounds", [
    ("test1.vsdx", 0, "Shape Text", ['0.25', '9.87', '2.42', '11.44']),  # standard shape
    ("test2.vsdx", 0, "Sub-shape 2", ['0.79', '10.04', '1.87', '10.37']),  # sub shape in a group
    ("test2.vsdx", 0, "Scenario:", ['0.73', '8.19', '2.49', '8.97']),  # line
    ])
def test_shape_relative_bounds(filename, page_index, shape_text, expected_bounds):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index]
        shape = page.find_shape_by_text(shape_text)
        print([f"{n:.2f}" for n in shape.relative_bounds])
        assert [f"{n:.2f}" for n in shape.relative_bounds] == expected_bounds


@pytest.mark.parametrize("filename, page_index, shape_text, arrow", [
    ("test2.vsdx", 0, "Scenario:", True),  # add end arrow
    ("test2.vsdx", 0, "Scenario:", False),  # no end arrow
    ])
def test_shape_end_arrow(filename, page_index, shape_text, arrow):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_shape_end_arrow_{arrow}.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index]
        shape = page.find_shape_by_text(shape_text)
        shape.end_arrow = arrow
        print(shape.end_arrow)
        print(vsdx.pretty_print_element(shape.xml))
        assert shape.end_arrow == '13' if arrow else '0'
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.pages[page_index]
        shape = page.find_shape_by_text(shape_text)
        print(shape.end_arrow)
        print(vsdx.pretty_print_element(shape.xml))
        assert shape.end_arrow == '13' if arrow else '0'


@pytest.mark.parametrize("filename, page_index, shape_text, expected_universal_name", [
    ("test3_house.vsdx", 0, "context filter", "House"),
    ("test4_connectors.vsdx", 1, "Shape A", None),  # no master and no name
    ("test4_connectors.vsdx", 1, "A to B", "Dynamic connector"),  # master with Layer and UnivName Cell
    ("test4_connectors.vsdx", 2, "Switch", "Switch"),  # master with no Layer
    ("test4_connectors.vsdx", 2, "Router", "Router"),  # master with no Layer
    ])
def test_shape_universal_name(filename, page_index, shape_text, expected_universal_name):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index]
        shape = page.find_shape_by_text(shape_text)
        assert shape.universal_name == expected_universal_name


@pytest.mark.parametrize("filename, expected_master_shape_name",
                         [("test_master_multiple_child_shapes.vsdx",  # master with multiple child shapes
                           "AWS Step Functions workflow ")])  # expected master shape name
def test_get_shape_master_page(filename: str, expected_master_shape_name: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        child_shape = vis.get_page(0).child_shapes[0]
        assert child_shape.master_page.name == expected_master_shape_name


@pytest.mark.parametrize("filename, shape_id, expected_angle",
                         [
                             ("test11_rotate.vsdx", "1", 0.52),
                             ("test11_rotate.vsdx", "2", -1.39),
                             ("test11_rotate.vsdx", "5", 2.53),
                             ("test11_rotate.vsdx", "6", -0.26),
                         ])
def test_get_shape_angle(filename: str, shape_id: str, expected_angle: float):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]
        # check angle is close enough to expected
        assert abs(page.find_shape_by_id(shape_id).angle - expected_angle) < 0.01


@pytest.mark.parametrize(("filename", "regex", "expected_shape_ids"),
                         [
                             ('test1.vsdx', r'\s(\S{2})\s', ['2', '5', '6']),
                             
                          ])
def test_find_shapes_by_regex(filename: str, regex: str, expected_shape_ids: list):
    """ Test function Shape.find_shapes_by_regex(regex: str) 
        test1.vsdx contains Shapes with text fields
            [(shp.ID,shp.text) for shp in shapes.all_shapes]:
            ('1', 'Shape Text\n')
            ('2', 'Shape to remove\n')
            ('5', 'Shape to copy\n')
            ('6', 'Shape for context filter: The scenario is {{scenario}} and this file was created on {{date}}\n')
        
    the regex matches the sequence '(whitespace)(two non-whitespace)(whitespace)' as for example with the strings ' to ' in ID=2,5 and also ' is ' and ' on ' in ID=6
    so we expect IDs=[2,5,6]
    """
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shapes = vis.pages[0].shapes[0]  # type: Page
        assert len(shapes.find_shapes_by_regex('')) == len(shapes.all_shapes)
        fil_shapes = shapes.find_shapes_by_regex(regex)
        assert [shp.ID for shp in fil_shapes] == expected_shape_ids


@pytest.mark.parametrize(("filename", "page_index", "shape_text", "color_param", "expected_colour"),
                         [
                             ('test12_colors.vsdx', 0, "Line Color", "line", "#ff0000"),
                             ('test12_colors.vsdx', 0, "Text Color", "text", "#ff0000"),
                             ('test12_colors.vsdx', 0, "Fill Color", "fill", "#ff0000"),
                             ('test12_colors.vsdx', 1, "Line Color", "line", "#00ff00"),
                             ('test12_colors.vsdx', 1, "Text Color", "text", "#00ff00"),
                             ('test12_colors.vsdx', 1, "Fill Color", "fill", "#00ff00"),

                         ])
def test_get_shape_line_color(filename: str, page_index: int, shape_text: str, color_param: str, expected_colour: str):
    """Test that we can get a shapes line, text, or fill color"""
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shape = vis.pages[page_index].find_shape_by_text(shape_text)
        print(f"LineColor={shape.line_color} FillColor={shape.fill_color} TextColor={shape.text_color}")
        if color_param == "line":
            assert shape.line_color == expected_colour
        elif color_param == "text":
            assert shape.text_color == expected_colour
        elif color_param == "fill":
            assert shape.fill_color == expected_colour


@pytest.mark.parametrize(("filename", "page_index", "shape_text", "color_param", "expected_colour"),
                         [
                             ('test12_colors.vsdx', 0, "Line Color", "line", "#00ff00"),
                             ('test12_colors.vsdx', 0, "Text Color", "text", "#00ff00"),
                             ('test12_colors.vsdx', 0, "Fill Color", "fill", "#00ff00"),
                             ('test12_colors.vsdx', 1, "Line Color", "line", "#0000ff"),
                             ('test12_colors.vsdx', 1, "Text Color", "text", "#0000ff"),
                             ('test12_colors.vsdx', 1, "Fill Color", "fill", "#0000ff"),

                         ])
def test_set_shape_line_color(filename: str, page_index: int, shape_text: str, color_param: str, expected_colour: str):
    """Test that we can set a shapes line, text, or fill color"""
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_{page_index}_set_shape_{color_param}_color.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        shape = vis.pages[page_index].find_shape_by_text(shape_text)
        print(f"LineColor={shape.line_color} FillColor={shape.fill_color} TextColor={shape.text_color}")
        if color_param == "line":
           shape.line_color = expected_colour
        elif color_param == "text":
            shape.text_color = expected_colour
        elif color_param == "fill":
            shape.fill_color = expected_colour
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        shape = vis.pages[page_index].find_shape_by_text(shape_text)
        print(f"LineColor={shape.line_color} FillColor={shape.fill_color} TextColor={shape.text_color}")
        if color_param == "line":
            assert shape.line_color == expected_colour
        elif color_param == "text":
            assert shape.text_color == expected_colour
        elif color_param == "fill":
            assert shape.fill_color == expected_colour
