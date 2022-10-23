.. vsdx documentation master file, created by
   sphinx-quickstart on Sat May  8 21:14:08 2021.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to vsdx's documentation!
================================
v0.5.11

.. toctree::
   :maxdepth: 2
   :caption: Contents:

   quickstart
   find_shape
   classes

* :ref:`genindex`

Project
-------

vsdx is designed to allow a developer to programmatically alter .vsdx (Visio) diagram files

Features
--------

The vsdx package provides a VisioFile class, which can instantiated using an existing vsdx file as below:

.. code-block:: python

   from vsdx import VisioFile  # import the package

   with VisioFile('diagram.vsdx') as vis:  # open file using context manager
      pass  # do nothing yet

The VisioFile object can then be used to access objects such as Pages, Shapes, and Connects.

There are two models of interaction with the VisioFile object.

* Jinja
* Attribute manipulation

Jinja
-----

Jinja https://pypi.org/project/Jinja2/ is a great templating package.

vsdx allow you to create vsdx files from vsdx templates, with just a few lines of code...

.. code-block:: python

   from vsdx import VisioFile  # import the package

   with VisioFile('template_diagram.vsdx') as vis:  # open file using context manager
      context = {'author':'Dave', 'numbers':[1,2,3] }
      vis.jinja_render_vsdx(context=context)
      vis.save_vsdx('my_new_file.vsdx')

... it's as simple as that. The data passed in the 'context' dictionary is available to the Jinja template.
For example `{{ author }}` would result in 'dave', and `{{ author|lower }}` would result in 'dave'.


Attribute Manipulation
----------------------

The VisioFile object provides access to a list of VisioFile.Page objects

Shape's can be found in a Page either by ID (if known) or by the text they contain.

Various methods can then be applied to a VisioFile.Shape object - such as updating the shape text, position, or even
removing or copying the shape.

.. code-block:: python

   from vsdx import VisioFile  # import the package

   with VisioFile('diagram.vsdx') as vis:  # open file using context manager
      page = vis.page_objects[0]  # type: VisioFile.Page
      shape = page.find_shape_by_text('the shape you are looking for')  # type: VisioFile.Shape
      shape.text = "some new text for the shape"  # update the text
      shape.x += 1.0  # move the shape to where you want it

