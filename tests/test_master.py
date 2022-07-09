import pytest
import os

import vsdx
from vsdx import Page  # for typing
from vsdx import Shape  # for typing
from vsdx import VisioFile

# code to get basedir of this test file in either linux/windows
basedir = os.path.dirname(os.path.relpath(__file__))


@pytest.mark.parametrize(("filename", "expected_length"),
                         [('test5_master.vsdx', 1)])
def test_load_master_file(filename: str, expected_length: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        assert len(vis.master_pages) == expected_length


@pytest.mark.parametrize(("filename", "shape_text"),
                         [('test5_master.vsdx', "Shape B")])
def test_find_master_shape(filename: str, shape_text: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        master_page =  vis.master_pages[0]  # type: Page
        s = master_page.find_shape_by_text(shape_text)
        assert s


@pytest.mark.parametrize(("filename"),
                         [('test5_master.vsdx')])
def test_master_inheritance(filename: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        master_page = vis.master_pages[0]  # type: Page
        shape_a = page.find_shape_by_text('Shape A')  # type: Shape
        shape_b = page.find_shape_by_text('Shape B')  # type: Shape

        assert shape_a
        assert shape_b

        # test inheritance of cell values
        assert shape_a.width == 1
        assert shape_a.height == 0.75

        # test inheritance for subshapes
        sub_shape_a = shape_a.child_shapes[0]
        assert sub_shape_a.cell_value('LineWeight') == '0.01875'


@pytest.mark.parametrize(("filename"),
                         [('test5_master.vsdx')])
def test_set_master_child_property(filename: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_set_master_child_property.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)
        # find shape with master and set a child property
        sub_shape_b = page.find_shape_by_text('Shape B').child_shapes[0]

        sub_shape_b.line_weight = 0.5

        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page(0)

        sub_shape_b = page.find_shape_by_text('Shape B').child_shapes[0]
        assert sub_shape_b.master_shape  # shape has a master
        assert sub_shape_b.line_weight == 0.5  # child shape has value set


@pytest.mark.parametrize(("filename, weight"),
                         [('test5_master.vsdx', 0.5)])
def test_master_property_change_is_inherited(filename: str, weight):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_master_property_change_is_inherited.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)

        sub_shape_a = page.find_shape_by_text('Shape A').child_shapes[0]
        master = sub_shape_a.master_shape
        master.line_weight = weight

        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page(0)

        sub_shape_a = page.find_shape_by_text('Shape A').child_shapes[0]
        sub_shape_b = page.find_shape_by_text('Shape B').child_shapes[0]

        # check that both sub_shape values have changed based on master
        assert sub_shape_a.line_weight == weight
        assert sub_shape_b.line_weight == weight


@pytest.mark.parametrize(("filename, weight"),
                         [('test5_master.vsdx', 0.1)])
def test_child_property_change_is_not_inherited(filename: str, weight):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_child_property_change_is_not_inherited.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)

        shape_a = page.find_shape_by_text('Shape A')
        sub_shape_a = shape_a.child_shapes[0]
        master = sub_shape_a.master_shape
        original_master_weight = master.line_weight
        # increment child property to ensure child and master are not the same
        line_weight = master.line_weight + weight
        sub_shape_a.line_weight = line_weight  # change the child, not the master shape

        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page(0)

        sub_shape_a = page.find_shape_by_text('Shape A').child_shapes[0]
        master = sub_shape_a.master_shape

        # check that child shape has changed to expected value
        assert sub_shape_a.line_weight == line_weight
        # check that master property has not changed
        assert master.line_weight == original_master_weight


@pytest.mark.parametrize(("filename", "shape_text"),
                         [("test_master.vsdx", "Page Shape"),
                          ("test_master.vsdx", "Master Shape A"),
                          ("test_master.vsdx", "Master Shape B"),
                          ("test_master.vsdx", "Master B with updated text"), ])
def test_master_find_shapes(filename: str, shape_text: str):
    # Check that shape with text can be found - whether in page, master or overridden in master
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
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
        page = vis.get_page(0)  # type: Page
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
