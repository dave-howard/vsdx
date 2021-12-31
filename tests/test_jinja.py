import pytest
from vsdx import VisioFile, namespace, vt_namespace, ext_prop_namespace, PagePosition, Media
from datetime import datetime
import os
from typing import List

# code to get basedir of this test file in either linux/windows
basedir = os.path.dirname(os.path.relpath(__file__))

@pytest.mark.parametrize(("filename", "context"),
                         [("test_jinja.vsdx", {"date": datetime.now(), "scenario": "Scenario One", "x": 2, "y": 2}),
                          ("test_jinja.vsdx", {"date": datetime.now(), "scenario": "Scenario Two", "x": 2, "y": 2}),
                          ("test_jinja.vsdx", {"date": datetime.now(), "scenario": "Scenario Three", "x": 2, "y": 2}),
                          ])
def test_basic_jinja(filename: str, context: dict):

    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_basic_jinja.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]

        # each key string in context dict will be replaced with value, record against shape Ids for validation
        shape_id_values = dict()
        for k, v in context.items():
            shape_id = page.find_shape_by_text(k).ID
            shape_id_values[shape_id] = v
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and validate each shape id has expected text
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        for shape_id, text in shape_id_values.items():
            if type(text) is str:
                print(f"Testing that shape {shape_id} has text '{text}' in: {page.find_shape_by_id(shape_id).text}")
                assert str(text) in page.find_shape_by_id(shape_id).text


@pytest.mark.parametrize(("filename", "context", "shape_count"),
                         [("test_jinja.vsdx", {"date": datetime.now(), "scenario": "One", "x": 0, "y": 2}, 1),
                          ("test_jinja.vsdx", {"date": datetime.now(), "scenario": "Two", "x": 5, "y": 2}, 2),
                          ("test_jinja.vsdx", {"date": datetime.now(), "scenario": "Three", "x": 20, "y": 2}, 3),
                          ])
def test_jinja_if(filename: str, context: dict, shape_count: int):

    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_jinja_if_{context["scenario"]}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and validate each shape id has expected text
    with VisioFile(out_file) as vis:
        page = vis.pages[1]  # second page has the shapes with if statements
        count = len(page.sub_shapes())
        print(f"expected {shape_count} and found {count}")
        assert count == shape_count


@pytest.mark.parametrize(("filename", "context"),
                         [("test_jinja.vsdx", {"x": 2, "y": 2}),
                          ("test_jinja.vsdx", {"x": 3, "y": 4}),
                          ("test_jinja.vsdx", {"x": 12, "y": 9}),
                          ])
def test_jinja_calc(filename: str, context: dict):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_jinja_calc.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and validate each shape id has expected text
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        # check a shape exists with product of x and y values
        x_y = str(context['x'] * context['y'])
        assert page.find_shape_by_text(x_y)


@pytest.mark.parametrize(("filename", "context"),
                         [("test_jinja_loop.vsdx", {"date": datetime.now(), "scenario": "Scenario One", "test_list":[1,2,3]}),
                          ("test_jinja_loop.vsdx", {"date": datetime.now(), "scenario": "Scenario Two", "test_list":["One", "Two","Three"]}),
                          ("test_jinja_loop.vsdx", {"date": datetime.now(), "scenario": "Scenario Three", "test_list":[1,2,3,4,5,6]}),
                          ])
def test_basic_jinja_loop(filename: str, context: dict):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_basic_jinja.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]

        # each key string in context dict will be replaced with value, record against shape Ids for validation
        shape_id_values = dict()
        for k, v in context.items():
            shape_id = page.find_shape_by_text(k).ID
            shape_id_values[shape_id] = v
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and validate each shape id has expected text, and that a shape exists with each loop value
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        for shape_id, text in shape_id_values.items():
            if type(text) is str:
                print(f"Testing that shape {shape_id} has text '{text}' in: {page.find_shape_by_id(shape_id).text}")
                assert str(text) in page.find_shape_by_id(shape_id).text
            if type(text) is list:
                for item in text:
                    print(f"Testing that shape with text '{item}' exists")
                    assert page.find_shape_by_text(str(item))


@pytest.mark.parametrize(("filename", "context"),
                         [("test_jinja_inner_loop.vsdx", {"test_list":[[1, 2, 3], [1, 2, 3], [1, 2, 3]]}),
                          ("test_jinja_inner_loop.vsdx", {"test_list":["One", "Two","Three"]}),
                          ])
def test_jinja_inner_loop(filename: str, context: dict):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_jinja_inner_loop.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and validate each shape id has expected text, and that a shape exists with each loop value
    test_list = context["test_list"]
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        for o in test_list:
            for p in o:
                print(f"p={p}")
                s = page.find_shape_by_text(str(p))
                assert s


@pytest.mark.parametrize(("filename", "out_name", "context"),
                         [("test_jinja_loop_showif.vsdx", "1234", {"test_list": [1, 2, 3, 4]}),
                          ("test_jinja_loop_showif.vsdx", "3456", {"test_list": [3, 4, 5, 6]}),
                          ])
def test_jinja_loop_showif(filename: str, out_name: str,  context: dict):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}{out_name}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        for o in context['test_list']:
            s = page.find_shape_by_text(f"In this instance, o={o}")
            # test that {% show if o > 2 %} has worked
            print(f"for {o} found {s}")
            if o > 2:
                assert s  # check that shape is found (not removed by showif)
            else:
                assert not s  # check that shape is not found (has been removed by showif)


@pytest.mark.parametrize(
    ("filename", "context", "shape_id", "expected_x", "expected_text"),
    [("test_jinja_self_refs.vsdx", {"n": 1}, "1", 2.0, "This text should remain  and x should be 2.0\n"),
     ("test_jinja_self_refs.vsdx", {"n": 2}, "2", 4.0, "This shape sets x to n * 2\n"),
     ("test_jinja_self_refs.vsdx", {"n": 1}, "3", 1.0, "This shape sets x to 1 if n else 2\n"),
     ("test_jinja_self_refs.vsdx", {"n": 0}, "3", 2.0, "This shape sets x to 1 if n else 2\n"),
     ])
def test_jinja_self_refs(filename: str, context: dict, shape_id, expected_x, expected_text):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_jinja_self_refs.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: VisioFile.Page
        # there should be one shape on page 0
        shape = page.find_shape_by_id(shape_id)  # type: VisioFile.Shape
        print(f"DEBUG: ID={shape.ID} shape.text={shape.text}")
        print(f"DEBUG: ID={shape.ID} shape.x={shape.x}")
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and check shape has moved
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        shape = page.find_shape_by_id(shape_id)  # type: VisioFile.Shape
        print(f"DEBUG: ID={shape.ID} shape.text='{shape.text}' expected='{expected_text}'")
        print(f"DEBUG: ID={shape.ID} shape.x={shape.x}")
        assert shape.x == expected_x
        assert shape.text == expected_text


@pytest.mark.parametrize(
    ("filename", "context", "shape_id", "expected_y", "expected_text"),
    [("test_jinja_self_refs.vsdx", {"n": 2}, "4", 10.12368731806121, "This shape should move down by 1.0\n"),
     ("test_jinja_self_refs.vsdx", {"n": 1}, "5", 8.726049539918966, "This shape should move down by n\n"),
     ("test_jinja_self_refs.vsdx", {"n": 2}, "5", 7.726049539918966, "This shape should move down by n\n"),
     ])
def test_jinja_self_ref_calculations(filename: str, context: dict, shape_id, expected_y, expected_text):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_jinja_self_ref_calcs.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: VisioFile.Page
        # there should be one shape on page 0
        shape = page.find_shape_by_id(shape_id)  # type: VisioFile.Shape
        if shape:
            print(f"DEBUG: ID={shape.ID} shape.text={shape.text}")
            print(f"DEBUG: ID={shape.ID} shape.y={shape.y}")
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and check shape has moved
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        shape = page.find_shape_by_id(shape_id)  # type: VisioFile.Shape
        print(f"DEBUG: ID={shape.ID} shape.text='{shape.text}' expected='{expected_text}'")
        print(f"DEBUG: ID={shape.ID} shape.x={shape.y}")
        assert shape.y == expected_y
        assert shape.text == expected_text


@pytest.mark.parametrize(
    ("filename", "context", "expected_page_count", "expected_page_names"),
    [("test_jinja_page_showif.vsdx", {"show": True}, 2, ['Normal Page', 'Page2']),
     ("test_jinja_page_showif.vsdx", {"show": 1}, 2, ['Normal Page', 'Page2']),
     ("test_jinja_page_showif.vsdx", {"show": "true"}, 2, ['Normal Page', 'Page2']),
     ("test_jinja_page_showif.vsdx", {"show": 4.0}, 2, ['Normal Page', 'Page2']),
     ("test_jinja_page_showif.vsdx", {"show": False}, 2, ['Normal Page', 'Page3']),
     ("test_jinja_page_showif.vsdx", {"show": []}, 2, ['Normal Page', 'Page3']),
     ("test_jinja_page_showif.vsdx", {"show": {}}, 2, ['Normal Page', 'Page3']),
     ])
def test_jinja_page_showif(filename: str, context: dict, expected_page_count, expected_page_names):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_show_{context["show"]}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        print(f"len(vis.pages)={len(vis.pages)} context={context}")
        for p in vis.pages: # type: VisioFile.Page
            print(f"page:{p.name}")
        vis.jinja_render_vsdx(context=context)
        vis.save_vsdx(out_file)

    # open file and check shape has moved
    with VisioFile(out_file) as vis:
        page_names = []
        for p in vis.pages: # type: VisioFile.Page
            print(f"page:{p.name}")
            page_names.append(p.name)
        assert len(vis.pages) == expected_page_count
        assert page_names == expected_page_names
