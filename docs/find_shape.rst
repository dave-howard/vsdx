Finding Shapes
==============

The anatomy of a vsdx file is complex - but can be summarised as Pages, containing Shapes, which might contain their own child Shapes (and so on).

When using vsdx to process a diagram it is likely that you will want to act on a specific shape or shapes.
To do this you will typically first identify the page, and then use some criteria to select the shape or shapes.

Selecting a Page
----------------

There are two ways to select a Page - either by (zero-based) index, or by (case-sensitive) name.

**Select a Page by index**

.. code-block:: python

    from vsdx import VisioFile  # import the package

    with VisioFile('diagram.vsdx') as vis:  # create a VisioFile object from a file
        page = vis.pages[0]  # get the first Page in vsdx file
        print(page.name)  # print the name of the page


**Select a Page by name**

.. code-block:: python

    from vsdx import VisioFile  # import the package

    with VisioFile('diagram.vsdx') as vis:  # create a VisioFile object from a file
        page = vis.get_page_by_name('Page 1')  # get Page with name 'Page 1
        print(page.name)  # print the name of the page


Selecting a Shape or Shapes
-----------------

**Find Shape in a Page**

You can select a Shape or list of Shapes by ID, shape text, shape property

.. code-block:: python

    from vsdx import VisioFile  # import the package

    with VisioFile('diagram.vsdx') as vis:  # create a VisioFile object from a file
        page = vis.pages[0]  # get first page
        shape_1 = page.find_shape_by_id('1')  # get Shape in page with ID of '1'
        shape_A = page.find_shape_by_text('A')  # get first Shape in page where text contains 'A'
        shape_with_red_label = page.find_shape_by_property_label('red')  # get first Shape in page with a property label of 'red'
        code_a_shape = page.find_shape_by_property_label_value('code', 'a')  # get first Shape in page where property 'code' = 'a'

Each of these 'find_shape_' methods returns a Shape object (or None).

**Note:** the Shape property label is the one that is visible in the 'Shape Data Window' in Microsoft Visio, it is not the same as the property name.

**Find Shapes in a Page**

Sometimes you want to find a group of related Shapes in a Page.

**Note:** when selecting by ID, as the ID should be unique, no find_shapes_by_id() method is available.

.. code-block:: python

    from vsdx import VisioFile  # import the package

    with VisioFile('diagram.vsdx') as vis:  # create a VisioFile object from a file
        page = vis.pages[0]  # get first page
        shapes_A = page.find_shapes_by_text('A)  # get all Shapes in page where text contains 'A'
        shapes_with_red_label = page.find_shapes_by_property_label('red')  # get all Shapes in page with a property label of 'red'
        code_a_shapes = page.find_shapes_by_property_label_value('code', 'a')  # get all Shapes in page where property 'code' = 'a'

Each of these 'find_shapes_' method returns typed list of Shapes - List[Shape]

Selecting within a Shape
------------------------

As a Shape can contain other Shapes, which themselves can contain more Shapes. Each Shape object supports the same methods as Page for selecting shapes, as per exampels below:

.. code-block:: python

    from vsdx import VisioFile  # import the package

    with VisioFile('diagram.vsdx') as vis:  # create a VisioFile object from a file
        page = vis.pages[0]  # get first page
        group_shape = page.find_shape_by_property_label('my_container_1')
        my_shape = group_shape.find_by_text('my shape')  # find a Shape within the group shape

In the example above, there might be many Shapes with the text 'my shape', but this will find the first within your group shape tagged with a property/label

Selecting all or child Shapes
-------------------
As mentioned above, each Page is a hierarchy of Shapes.

Both Page and Shape provide all_shapes and child_shapes properties.

.. code-block:: python

    from vsdx import VisioFile  # import the package

    with VisioFile('diagram.vsdx') as vis:
        # open first page
        page = vis.pages[0]
        # Page.child_shapes and Page.all_shapes properties
        page_top_shapes = page.child_shapes  # just those Shapes directly under Page
        all_shapes_in_page = page.all_shapes  # all shapes in the hierarchy of Page
        # Shape.child_shapes and Shape.all_shapes properties
        shape = page.all_shapes[0] #  get first Shape in page
        shape_children = shape.child_shapes  # just Shapes directly under Shape
        shape_all_shapes = shape.all_shapes  # all children in hierarchy under Shape


