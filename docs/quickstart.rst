Quick Start Guide
=================

You can start using vdsx in just a few lines of code!

Installation
------------

vsdx is a python package - so like any package you can install from the terminal along with any missing dependencies using pip.

``pip install vsdx``

**Note:** You will need `Python 3.7+` to use vsdx

Opening your first vsdx file
----------------------------

Your first two lines of code will open a vsdx file.

.. code-block:: python

   from vsdx import VisioFile  # import the package

   vis = VisioFile('diagram.vsdx')  # create a VisioFile object from a file


Making some changes
-------------------
You might want to open a page, find a shape with the text 'foo' and replace it with 'bar'
but first - let's use a context manager to make sure the file is closed when we are done

.. code-block:: python

    from vsdx import VisioFile  # import the package

    with VisioFile('diagram.vsdx') as vis:
        # open first page
        page = vis.pages_objects[0]  # type: VisioFile.Page
        # find a shape by text
        shape = page.find_shape_by_text('foo')  # type: VisioFile.Shape
        shape.text = 'bar'


Saving a copy
-------------
You'll want save your changes - lets just add that one line

.. code-block:: python

    from vsdx import VisioFile  # import the package

    with VisioFile('diagram.vsdx') as vis:
        # open first page
        page = vis.pages_objects[0]  # type: VisioFile.Page
        # find a shape by text
        shape = page.find_shape_by_text('foo')  # type: VisioFile.Shape
        shape.text = 'bar'
        vis.save_vsdx('copy_of_diagram.vsdx')  # save to a new file

and that's that.

For more detailed examples please have a look at the tests.py file
