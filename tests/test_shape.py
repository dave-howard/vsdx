import pytest
import os

import vsdx
from vsdx import DataProperty
from vsdx import Page  # for typing
from vsdx import Shape
from vsdx import VisioFile

# code to get basedir of this test file in either linux/windows
basedir = os.path.dirname(os.path.relpath(__file__))


@pytest.mark.parametrize("filename, expected_locations",
                         [("test1.vsdx", "1.33,10.66 4.13,10.66 6.94,10.66 2.33,9.02 "),
                          ("test2.vsdx", "2.33,8.72 1.33,10.66 4.13,10.66 5.91,8.72 1.61,8.58 3.25,8.65 ")])
def test_shape_locations(filename: str, expected_locations: str):
    print("=== list_shape_locations ===")
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        shapes = page.child_shapes
        locations = ""
        for s in shapes:  # type: Shape
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
        page = vis.get_page(0)  # type: Page
        shape = page.find_shape_by_id(shape_id)

        print(f"shape {shape.ID} center={shape.center_x_y}")
        print(f"x {shape.x} y={shape.y}")
        print(f"end_x {shape.end_x} end_y={shape.end_y}")
        assert shape.center_x_y == expected_center


@pytest.mark.parametrize("filename", ["test2.vsdx", "test3_house.vsdx"])
def test_remove_shape(filename: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_shape_removed.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        # get shape to remove
        s = vis.pages[0].find_shape_by_text('Shape to remove')  # type: Shape
        assert s  # check shape found
        s.remove()
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        # get shape that should have been removed
        s = vis.pages[0].find_shape_by_text('Shape to remove')  # type: Shape
        assert s is None  # check shape not found


@pytest.mark.parametrize(("filename", "shape_names", "shape_locations"),
                         [("test1.vsdx", {"Shape to remove"}, {(1.0, 1.0)})])
def test_set_shape_location(filename: str, shape_names: set, shape_locations: set):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_set_shape_location.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        # move shapes in list
        for (shape_name, x_y) in zip(shape_names, shape_locations):
            s = vis.pages[0].find_shape_by_text(shape_name)  # type: Shape
            assert s  # check shape found
            assert s.x
            assert s.y
            print(f"Moving shape '{shape_name}' from {s.x}, {s.y} to {x_y}")
            s.x = x_y[0]
            s.y = x_y[1]
        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        # get each shape that should have been moved
        for (shape_name, x_y) in zip(shape_names, shape_locations):
            s = vis.pages[0].find_shape_by_text(shape_name)  # type: Shape
            assert s  # check shape found
            assert s.x == x_y[0]
            assert s.y == x_y[1]


@pytest.mark.parametrize(("filename", "page_index", "shape_names", "shape_x_y_deltas"),
                         [("test1.vsdx", 0, {"Shape to remove"}, {(1.0, 1.0)}),
                          ("test4_connectors.vsdx", 0, {"Shape B"}, {(1.0, 1.0)}),
                          ("test4_connectors.vsdx", 1, {"B to C"}, {(1.0, 1.0)})])
def test_move_shape(filename: str, page_index: int, shape_names: set, shape_x_y_deltas: set):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_move_shape.vsdx')
    expected_shape_locations = dict()

    with VisioFile(os.path.join(basedir, filename)) as vis:
        # move shapes in list
        for (shape_name, x_y) in zip(shape_names, shape_x_y_deltas):
            s = vis.pages[page_index].find_shape_by_text(shape_name)  # type: Shape
            assert s  # check shape found
            assert s.x
            assert s.y
            expected_shape_locations[shape_name] = (s.x + x_y[0], s.y + x_y[1])
            s.move(x_y[0], x_y[1])

        vis.save_vsdx(out_file)
    print(f"expected_shape_locations={expected_shape_locations}")

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        # get each shape that should have been moved
        for shape_name in shape_names:
            s = vis.pages[page_index].find_shape_by_text(shape_name)  # type: Shape
            assert s  # check shape found
            assert s.x == expected_shape_locations[shape_name][0]
            assert s.y == expected_shape_locations[shape_name][1]


@pytest.mark.parametrize(("filename", "shape_name"),
                         [("test1.vsdx", "Shape to copy"),
                          ("test4_connectors.vsdx", "Shape B")])
def test_shape_copy(filename: str, shape_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_shape_copy.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page =  vis.pages[0]  # type: Page
        # find and copy shape by name
        s = page.find_shape_by_text(shape_name)  # type: Shape
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
        page =  vis.pages[0]  # type: Page
        page2 = vis.pages[1]  # type: Page
        page3 = vis.pages[2]  # type: Page
        # find and copy shape by name
        s = page.find_shape_by_text(shape_name)  # type: Shape
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
        shape = vis.pages[0].find_shape_by_text(shape_name)  # type: Shape

        props = shape.data_properties

        # check overall type
        assert type(props) is dict

        # check each item type
        for prop in props.values():
            assert type(prop) is DataProperty


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
        shape = vis.pages[page_index].find_shape_by_text(shape_name)  # type: Shape

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