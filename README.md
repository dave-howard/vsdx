## vsdx - A python library for processing Visio .vsdx files

![PyPI](https://img.shields.io/pypi/v/vsdx)
![PyPI - Downloads](https://img.shields.io/pypi/dm/vsdx)

[![pytest](https://github.com/dave-howard/vsdx/actions/workflows/test.yaml/badge.svg)](https://github.com/dave-howard/vsdx/actions/workflows/test.yaml)
[![Documentation Status](https://readthedocs.org/projects/vsdx/badge/?version=latest)](https://vsdx.readthedocs.io/en/latest/?badge=latest)

![PyPI - Python Version](https://img.shields.io/pypi/pyversions/vsdx)
[![vsdx](https://snyk.io/advisor/python/vsdx/badge.svg)](https://snyk.io/advisor/python/vsdx)

__.vsdx files can be processed in two ways, directly via python code as in
example 1 below, or indirectly using a jinja template as in example 2__

For quick start documentation please see
[https://vsdx.readthedocs.io/en/latest/quickstart.html](https://vsdx.readthedocs.io/en/latest/quickstart.html)

__Example 1__ code to find a shape with specific text, remove it, and
then save the updated .vsdx file:

```python
from vsdx import VisioFile

filename = 'my_file.vsdx'
# open a visio file
with VisioFile(filename) as vis:
  # find shape by its text on first page
  shape = vis.pages[0].find_shape_by_text('Shape to remove')
  # remove the shape if found
  if shape:
    shape.remove()
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

----

###  Change Log
- v0.5.19: Fix Page.is_master_page prop and add tests
- v0.5.18: Add support for reading (not creating/saving) .vsdm files (macro-enabled)
- v0.5.17: Fix bug with adding multiple connectors to a vsdx with no existing master files/rels
- v0.5.16: Add support & commit hook tests for Python 3.11 and 3.12
- v0.5.15: Change DataProperty.value from field to property (get/set)
- v0.5.14: Add Page.find_shape_by_attr()
- v0.5.13: Update DataProperty class to get value of a property from V attrib or text
- v0.5.12: Add `Shape.fill_color` and `Shape.text_color` properties with get and set tests
- v0.5.11: Add `Shape.find_shapes_by_regex()` & `Page.find_shapes_by_regex()` - add check in `save_vsdx()` that file is open with more meaningful `VisioFileNotOpen` error
- v0.5.10: Add Shape.angle property
- v0.5.9: Add tests for master shape text property
- v0.5.8: Add `Page.master_base_id` property
- v0.5.7: Add support for nested shapes in `Page.all_shapes` and `Shape.all_shapes`. Add `Page.is_master_page` and `Shape.is_master_shape`. 
- v0.5.6: Fix error in `Shape.text` with missing master shape. Improve `VisioFile.remove_page_by_index()` and add `VisioFile.remove_page_by_name()`.
- v0.5.5: Added Shape.universal_name and used in Shape.set_start_and_finish(). Set Page.page_id on open file.
- v0.5.4: Added better (but still incomplete) support for adding connectors between shapes
- v0.5.3: Fixed missing deprecation dependency in setup.py
- v0.5.2: deprecated Page.set_name() method and page.page_name property, in favour of Page.name. Unskipped test: test_shape_center(). Added find_shape.rst doc page
- v0.5.1: added Page/Shape.find_shape_by_property_label_value()/find_shapes_by_property_label_value()
- v0.5.0: deprecated Page.shapes property Page/Shape.sub_shapes() methods in favour of Page/Shape.child_shapes property. Add Shape.all_shapes property, convert Page.all_shapes() method to property
- v0.4.20: Shape.set_cell_value()/set_cell_formula() create new cell if missing, add Media().rectangle and circle props, add Shape.bounds, relative_bounds, and end_arrow props
- v0.4.19: correctly position new connector shapes between 'from' and 'to' shapes 
- v0.4.18: add page.page_name, width and height properties
- v0.4.17: register xml namespaces to improve compatibility of xml output
- v0.4.16: Add Geometry class to read Shape geometry
- v0.4.15: Add requirement for Jinja2 to package
- v0.4.14: Include media/*.vsdx files in package
- v0.4.13: Fix for master shape property inheritance / value overrides
- v0.4.12: Add support for absolute paths
- v0.4.11: Add support for master shape data properties, with related tests
- v0.4.10: Add methods (`Shape.find_shape_by_property_label()` and `Shape.find_shapes_by_property_label`) to find shape or shapes by data property name.
- v0.4.9: Add support for creating new connection between two objects. Fix ShapeProperty.value, and add label, sort_key, value_type and prompt 
- v0.4.8: Support nested loops/showifs and combo of loop and if in same shape.
- v0.4.7: Python 3.10.0rc1 added to test suite. Add `Shape.data_properties` property, and new class `ShapeProperty` to represent Visio Shape Data
- v0.4.6: Add support for nested jinja loops in one shape
- v0.4.5: Fix bug where some shapes have no parent, and inserting shape into empty page
- v0.4.4: Added support for master page shape inheritance, ability to get `Shape.master_shape`, ability to 
  update master shapes and persist changes to master shapes in `save_vsdx()`
- v0.4.3: Added support for including/excluding pages via Jinja with `{% showif <statement> %}` in page name
- v0.4.2: Added `VisioFile.add_page_at()` method taking `index` to allow insertion
  at a specific point; Added `VisionFile.copy_page()` method to copy an existing page 
  and insert at a specific index or relative to copied page (using `PagePosition` enum). 
- v0.4.1: Added support for self referencing calculations in Jinja statements, 
  such as `{% set self.x = self x + n * 3.2 %}`
- v0.4.0: Added `VisioFile.jinja_set_selfs` to allow setting shape x and
  y properties in Jinja template. Setting values, calculations, or if
  statements are supported e.g. `{% set self.x = 1.5 %}` or `{% set
  self.y = n * 3 %}` or `{% set self.x = 1.0 if n else 2.0 %}`
- v0.3.5: Added `VisioFile.add_page()` method and tests
- v0.3.4: Added `VisioFile.remove_page_by_index()` method to remove a
  page, with associated test
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
  `Shape.copy()`. Find and load master pages when file is opened, store
  in VisioFile.master_page_objects and .master_pages
- v0.2.7: add `Shape.copy()` method
- v0.2.6: added `Page.get_connectors_between()` to get zero or many
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

