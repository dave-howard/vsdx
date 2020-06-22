# vsdx - A python library for processing Visio .vsdx files

## Note: this is an early release with limited functionality

example:
from vsdx import VisioFile

# open a visio file
with VisioFile('my_file.vsdx') as vis:
    # get page shapes collection
    shapes = vis.page_objects[0].shapes
    # get shape to remove by its text value
    s = shapes[0].find_shape_by_text('Shape to remove')  # type: VisioFile.Shape
    # remove the shape if found
    if s:
        s.remove()
        # save a new copy
        vis.save_vsdx(filename[:-5] + '_shape_removed.vsdx')


Please refer to tests/test.py for usage examples
