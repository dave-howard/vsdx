## vsdx - A python library for processing Visio .vsdx files

![PyPI](https://img.shields.io/pypi/v/vsdx)
![PyPI - Downloads](https://img.shields.io/pypi/dm/vsdx)

[![pytest](https://github.com/dave-howard/vsdx/actions/workflows/test.yaml/badge.svg)](https://github.com/dave-howard/vsdx/actions/workflows/test.yaml)
[![Documentation Status](https://readthedocs.org/projects/vsdx/badge/?version=latest)](https://vsdx.readthedocs.io/en/latest/?badge=latest)

![PyPI - Python Version](https://img.shields.io/pypi/pyversions/vsdx)
[![vsdx](https://snyk.io/advisor/python/vsdx/badge.svg)](https://snyk.io/advisor/python/vsdx)

__.vsdx files can processed in two ways, directly via python code as in
example 1 below, or indirectly using a jinja template as in example 2__

__Example 1__ code to find a shape with specific text, remove it, and
then save the updated .vsdx file:
```python
from vsdx import VisioFile

filename = 'my_file.vsdx'
# open a visio file
with VisioFile(filename) as vis:
    # get page shapes collection
    shapes = vis.page_objects[0].shapes
    # get shape to remove by its text value
    s = shapes[0].find_shape_by_text('Shape to remove')  # type: VisioFile.Shape
    # remove the shape if found
    if s:
        s.remove()
        # save a new copy
        vis.save_vsdx('shape_removed.vsdx')
```

__Example 2__ creating a new vsdx file from a template and context data
using jinja.  
Note that as vsdx does not lend itself well to ordered statements like
`{% if something %}my content{% endif %}` or `{% for x in list_value
%}x={{ x }}{% endfor %}` this package provides mechanisms to help -
refer to tests for more details.

```python
from vsdx import VisioFile

filename = 'my_template_file.vsdx'  # file containing jinja code
context = {'value1': 10, 'list_value': [1,2,3]}  # data for the template
with VisioFile('my_template_file.vsdx') as vis: 
    vis.jinja_render_vsdx(context=context)
    vis.save_vsdx('my_new_file.vsdx')
```

Please refer to tests/test.py for more usage
examples in the form of pytest tests.

---

###  Change Log
- v0.3.3: Added code of conduct and contributing guides
- v0.3.2: updated README and updated tests for improved compatibility
- v0.3.1: add jinja rendering support for if statements, via
  `VisioFile.jinja_render_vsdx()` - similar to for loops but using a `{%
  showif statement %}` in text of group shape controls whether that
  group shape is included in vsdx file rendered. Note that the showif
  statement is replaced with a standard if statement around the group
  shape prior to rendering. Refer to test.py::test_jinja_if() for an
  example
- v0.3.0: update jinja rendering to support for loops, where for
  statement is at start of group shape text, endfor is automatically
  inserted before processing. Refer to test.py::test_basic_jinja_loop()
  for code and test_jinja_loop.vsdx for content.
- v0.2.10: add `VisioFile.jinja_render_vsdx()` - applying jinja
  processing to Shape.text only
- v0.2.9: check that Page has shapes tag when in shape_copy(), add test
  to copy shape to new page
- v0.2.8: find max shape ID in Page before creating Shape in
  Shape.copy(). Find and load master pages when file is opened, store in
  VisioFile.master_page_objects and .master_pages
- v0.2.7: add Shape.copy() method
- v0.2.6: added Page.get_connectors_between() to get zero or many
  connectors between two shapes, by shape id or text
- 0.2.5: Add new Shape properties connected_shapes (list of Shape
  objects) and connects (list of Connect objects) properties to allow
  related shapes to be identified (i.e. shapes and connectors) and
  provide information on the relationship, in new Connect object. Also
  new properties of shape begin_x/y, end_x/y, plus height/width
  setters
- 0.2.4: Added find_replace(old, new) method to Shape and Page classes
  to recursively replace old with new
- 0.2.3: Updated tests to output files to an /out folder. Added test
  vsdx file with compound shape. Updated Shape text getter/setter
- 0.2.2: Added x & y location setters to Shape, and move(x_delta,
  y_delta) method - both with related tests

