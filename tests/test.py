from vsdx import VisioFile
from xml.etree.ElementTree import Element

# todo: convert to pytest test cases
# Simple tests - using test1.vsdx


def test_list_shape_id_and_text():
    vis = VisioFile('test1.vsdx')
    # loop through each page
    vis.open_vsdx_file()
    for file_name, page in vis.pages.items():
        shapes = vis.get_shapes(file_name)
        # loop though shapes in page
        for shape in shapes:  # type: Element
            text = vis.get_shape_text(shape)
            id = vis.get_shape_id(shape)
            print(f"Shape id:{id} text:{text}")
    vis.close_vsdx()


def test_set_shape_text():
    vis = VisioFile('test1.vsdx')
    # loop through each page
    vis.open_vsdx_file()
    for file_name, page in vis.pages.items():
        shapes = vis.get_shapes(file_name)
        # loop though shapes in page
        for shape in shapes:  # type: Element
            text = vis.get_shape_text(shape)
            if 'shape text' in text.lower():
                vis.set_shape_text(shape, 'Text Value Set')
    vis.save_vsdx('shape_one_text_changed.vsdx')


def test_remove_one_shape():
    vis = VisioFile('test1.vsdx')
    # loop through each page
    vis.open_vsdx_file()
    for file_name, page in vis.pages.items():
        shapes = vis.get_shapes(file_name)
        # loop though shapes in page
        remove_item = None
        for shape in shapes:  # type: Element
            text = vis.get_shape_text(shape)
            if 'shape to remove' in text.lower():
                remove_item = shape
        shapes.remove(remove_item)
    vis.save_vsdx('shape_two_removed.vsdx')


def test_copy_move_one_shape():
    vis = VisioFile('test1.vsdx')
    # loop through each page
    vis.open_vsdx_file()
    for file_name, page in vis.pages.items():
        shapes = vis.get_shapes(file_name)
        # loop though shapes in page
        copy_item = None
        for shape in shapes:  # type: Element
            text = vis.get_shape_text(shape)
            if 'shape to copy' in text.lower():
                copy_item = shape
        new_shape = vis.copy_shape(copy_item, page, file_name)
        vis.set_shape_text(new_shape, 'new copy of shape')
        if new_shape:  # move the shape
            x, y = vis.get_shape_location(new_shape)
            vis.set_shape_location(new_shape, x+1, y+1)

    vis.save_vsdx('shape_three_copied.vsdx')


if __name__ == '__main__':
    # run tests
    test_list_shape_id_and_text()   # prints shape id and text
    # these create three new updated copies of vsdx file
    test_set_shape_text()
    test_remove_one_shape()
    test_copy_move_one_shape()
