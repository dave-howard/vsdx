import pytest
import os

from vsdx import ext_prop_namespace
from vsdx import namespace
from vsdx import vt_namespace

from vsdx import Media
from vsdx import Page
from vsdx import PagePosition
from vsdx import Shape
from vsdx import VisioFile

# code to get basedir of this test file in either linux/windows
basedir = os.path.dirname(os.path.relpath(__file__))

# file structure

def test_invalid_file_type():
    filename = __file__
    print(f"Opening invalid but existing {filename}")
    try:
        with VisioFile(filename):
            assert False  # don't expect to get here
    except TypeError as e:
        print(e)  # expect to get here


def test_file_closure():
    filename = os.path.join(basedir, 'test1.vsdx')
    directory = f"./{filename.rsplit('.', 1)[0]}"
    with VisioFile(filename):
        # confirm directory exists
        assert os.path.exists(directory)
    # confirm directory is gone
    assert not os.path.exists(directory)


@pytest.mark.parametrize("filename", [
    "test1.vsdx",
    "diagram_with_macro.vsdm",
])
def test_open_rel_path(filename: str):
    # test opening a file in tests directory with relative path
    filename = os.path.join(basedir, filename)
    assert os.path.exists(filename)
    with VisioFile(filename) as vis:
        for page in vis.pages:
            print(page.name)


def test_open_abs_path():
    # test opening media file (not in tests directory)with absolute path
    media_file_path = Media()._media_vsdx.filename
    filename = os.path.abspath(media_file_path)

    assert(os.path.exists(filename))
    with VisioFile(filename) as vis:
        for p in vis.pages:
            print(p.name)


def test_open_abs_path_save_rel_path():
    # test opening media file (not in tests directory)with absolute path
    media_file_path = Media()._media_vsdx.filename
    filename = os.path.abspath(media_file_path)

    assert(os.path.exists(filename))
    output_file = os.path.join(basedir, 'out', 'abs_to_rel_out.vsdx')
    with VisioFile(filename) as vis:
        vis.save_vsdx(output_file)


def test_open_abs_path_save_abs_path():
    # test opening media file (not in tests directory)with absolute path
    media_file_path = Media()._media_vsdx.filename
    filename = os.path.abspath(media_file_path)

    assert(os.path.exists(filename))
    output_file = os.path.abspath(os.path.join(basedir, 'out', 'abs_to_abs_out.vsdx'))
    print('output_file', output_file)
    with VisioFile(filename) as vis:
        vis.save_vsdx(output_file)

# Helpers

@pytest.mark.parametrize(("filename", "shape_elements"), [("test1.vsdx", 4), ("test2.vsdx", 14), ("test3_house.vsdx", 10)])
def test_xml_findall_shapes(filename: str, shape_elements: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        # find all Shape elements
        xml = page.xml.getroot()
        xpath = f".//{namespace}Shape"
        elements = xml.findall(xpath)
        print(f"{xpath} returns {len(elements)} elements vs {shape_elements}")
        assert len(elements) == shape_elements


@pytest.mark.parametrize(("filename", "group_shape_elements"), [("test1.vsdx", 0), ("test2.vsdx", 3), ("test3_house.vsdx", 2)])
def test_xml_findall_group_shapes(filename: str, group_shape_elements: int):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        # find all Shape elements where attribute Type='Group'
        xml = page.xml.getroot()
        xpath = f".//{namespace}Shape[@Type='Group']"
        elements = xml.findall(xpath)
        print(f"{xpath} returns {len(elements)} elements vs {group_shape_elements}")
        assert len(elements) == group_shape_elements


# working with Pages

@pytest.mark.parametrize("filename, page_name", [("test1.vsdx", "Page-1"), ("test2.vsdx", "Page-1")])
def test_get_page(filename: str, page_name: str):
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.get_page(0)  # type: Page
        # confirm page name as expected
        assert page.name == page_name


@pytest.mark.parametrize(("filename"),
                         [("test1.vsdx"),
                          ])
def test_app_xml_page_names(filename: str):
    # test that page names in app.xml matches page names loaded
    with VisioFile(os.path.join(basedir, filename)) as vis:
        HeadingPairs = vis.app_xml.getroot().find(f'{ext_prop_namespace}HeadingPairs')
        i4 = HeadingPairs.find(f'.//{vt_namespace}i4')
        num_pages = int(i4.text)
        assert num_pages == len(vis.pages)

        # check page names from pages is same as page names from app.xml
        page_names = [p.name for p in vis.pages]
        TitlesOfParts = vis.app_xml.getroot().find(f'{ext_prop_namespace}TitlesOfParts')
        vector = TitlesOfParts.find(f'{vt_namespace}vector')
        app_xml_page_names = []
        for lpstr in vector.findall(f'{vt_namespace}lpstr'):
            page_name = lpstr.text
            app_xml_page_names.append(page_name)
        assert page_names == app_xml_page_names


@pytest.mark.parametrize(("filename", "page_index"),
                         [
                             ("test1.vsdx", 0),
                             ("test5_master.vsdx", 0),
                         ]
                         )
def test_remove_page_by_index(filename: str, page_index: int):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_remove_page.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page_count = len(vis.pages)
        vis.remove_page_by_index(page_index)
        vis.save_vsdx(out_file)

    # re-open file and confirm it has one less page
    with VisioFile(out_file) as vis:
        assert len(vis.pages) == page_count - 1


@pytest.mark.parametrize(("filename", "page_name"),
                         [
                             ("test1.vsdx", '1'),
                             ("test2.vsdx", '2'),
                             ("test3_house.vsdx", '1'),
                             ("test4_connectors.vsdx", '2'),
                             ("test5_master.vsdx", '1'),
                             ("test6_shape_properties.vsdx", '2'),
                             ("test7_with_connector.vsdx", '3'),
                             ("test8_simple_connector.vsdx", 'none'),
                         ])
def test_remove_page_by_page_index(filename: str, page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_remove_page_by_page_index.vsdx')
    expected_page_names = []
    with VisioFile(os.path.join(basedir, filename)) as vis:
        print(f"Found {len(vis.pages)} pages, {sorted([p.name for p in vis.pages])}")
        for page in list(vis.pages):
            print(f"page.name='{page.name}' page.page_name='{page.page_name}' index:{page.index_num}")
            if page_name in page.name:
                print(f"Removing page index {page.index_num}")
                vis.remove_page_by_index(page.index_num)
            else:
                expected_page_names.append(page.name)

        vis.save_vsdx(out_file)

    print(f"expected names={expected_page_names}")
    # re-open file and confirm it contains all and only those not deleted
    with VisioFile(out_file) as vis:
        print(f"Found {len(vis.pages)} pages, {sorted([p.name for p in vis.pages])}")
        assert sorted(expected_page_names) == sorted([p.name for p in vis.pages])


@pytest.mark.parametrize(("filename", "page_name"),
                         [
                             ("test1.vsdx", 'Page-1'),
                             ("test2.vsdx", 'Page-2'),
                             ("test3_house.vsdx", 'Page-1'),
                             ("test4_connectors.vsdx", 'Page-2'),
                             ("test5_master.vsdx", 'Page 1'),
                             ("test6_shape_properties.vsdx", 'Page-2'),
                             ("test7_with_connector.vsdx", 'Page-3'),
                             ("test8_simple_connector.vsdx", 'none'),  # no match
                         ])
def test_remove_page_by_name(filename: str, page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_remove_page_by_name.vsdx')
    expected_page_names = []
    with VisioFile(os.path.join(basedir, filename)) as vis:
        print(f"Found {len(vis.pages)} pages, {sorted([p.name for p in vis.pages])}")
        for page in list(vis.pages):
            print(f"page.name='{page.name}' page.page_name='{page.page_name}' index:{page.index_num}")
            if page_name != page.name:
                expected_page_names.append(page.name)
        vis.remove_page_by_name(page_name)
        vis.save_vsdx(out_file)

    print(f"expected names={expected_page_names}")
    # re-open file and confirm it contains all and only those not deleted
    with VisioFile(out_file) as vis:
        print(f"Found {len(vis.pages)} pages, {sorted([p.name for p in vis.pages])}")
        assert sorted(expected_page_names) == sorted([p.name for p in vis.pages])


@pytest.mark.parametrize(("filename", "remove_index"),
                         [("test1.vsdx", 0),
                          ("test1.vsdx", 1),
                          ("test2.vsdx", 0),
                          ])
def test_app_xml_page_names_after_remove_page(filename: str, remove_index: int):
    # test that page names in app.xml matches page names loaded
    with VisioFile(os.path.join(basedir, filename)) as vis:
        vis.remove_page_by_index(remove_index)

        HeadingPairs = vis.app_xml.getroot().find(f'{ext_prop_namespace}HeadingPairs')
        i4 = HeadingPairs.find(f'.//{vt_namespace}i4')
        num_pages = int(i4.text)
        assert num_pages == len(vis.pages)

        # check page names from pages is same as page names from app.xml
        page_names = [p.name for p in vis.pages]
        TitlesOfParts = vis.app_xml.getroot().find(f'{ext_prop_namespace}TitlesOfParts')
        vector = TitlesOfParts.find(f'{vt_namespace}vector')
        app_xml_page_names = []
        for lpstr in vector.findall(f'{vt_namespace}lpstr'):
            page_name = lpstr.text
            app_xml_page_names.append(page_name)
        assert page_names == app_xml_page_names


@pytest.mark.parametrize(('filename'),
                         [('test1.vsdx')])
def test_add_page(filename: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_add_page.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        number_pages = len(vis.pages)

        new_page = vis.add_page()
        assert new_page
        assert len(vis.pages) == number_pages + 1

        new_page_name = new_page.name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(new_page_name)
        assert page


@pytest.mark.parametrize(('filename', 'page_name'),
                         [
                            ('test1.vsdx', 'newname'),
                            ('test1.vsdx', 'Page-1'),
                         ])
def test_add_page_name(filename: str, page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_add_page_name_{page_name}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        number_pages = len(vis.pages)

        new_page = vis.add_page(page_name)
        assert new_page
        assert len(vis.pages) == number_pages + 1

        new_page_name = new_page.name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(new_page_name)
        assert page


@pytest.mark.parametrize(('filename', 'index', 'page_name'),
                         [
                            ('test1.vsdx', 1, None),
                            ('test1.vsdx', 1, 'newname'),
                            ('test1.vsdx', 1, 'Page-1'),
                         ])
def test_add_page_at(filename: str, index: int, page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_add_page_at_{page_name}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        number_pages = len(vis.pages)

        new_page = vis.add_page_at(index, page_name)
        assert new_page
        assert len(vis.pages) == number_pages + 1

        new_page_name = new_page.name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(new_page_name)
        assert page


@pytest.mark.parametrize(("filename", "new_page_name", "location"),
                         [("test1.vsdx", "new_page", 0),
                          ("test1.vsdx", "new_page", 1),
                          ("test2.vsdx", "new_page", 0),
                          ])
def test_app_xml_page_names_after_add_page(filename: str, new_page_name: str, location: int):
    # test that page names in app.xml matches page names loaded
    with VisioFile(os.path.join(basedir, filename)) as vis:
        if location is not None:
            vis.add_page(new_page_name)
        else:
            vis.add_page_at(location, new_page_name)

        HeadingPairs = vis.app_xml.getroot().find(f'{ext_prop_namespace}HeadingPairs')
        i4 = HeadingPairs.find(f'.//{vt_namespace}i4')
        num_pages = int(i4.text)
        assert num_pages == len(vis.pages)

        # check page names from pages is same as page names from app.xml
        page_names = [p.name for p in vis.pages]
        TitlesOfParts = vis.app_xml.getroot().find(f'{ext_prop_namespace}TitlesOfParts')
        vector = TitlesOfParts.find(f'{vt_namespace}vector')
        app_xml_page_names = []
        for lpstr in vector.findall(f'{vt_namespace}lpstr'):
            page_name = lpstr.text
            app_xml_page_names.append(page_name)
        assert page_names == app_xml_page_names


@pytest.mark.parametrize(('filename', 'index', 'page_name'),
                         [
                            ('test1.vsdx', -1, None),
                            ('test1.vsdx', 0, 'newname'),
                            ('test1.vsdx', 2, 'Page-1'),
                            ('test2.vsdx', 2, 'Page-1'),
                            ('test4_connectors.vsdx', 0, 'Page-1'),
                         ])
def test_copy_page(filename: str, index: int, page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_copy_page_{page_name}.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        number_pages = len(vis.pages)

        page = vis.pages[0]  # type: Page
        new_page = vis.copy_page(page, index=index, name=page_name)
        assert new_page
        assert len(vis.pages) == number_pages + 1

        new_page_name = new_page.name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(new_page_name)
        assert page


@pytest.mark.parametrize(('filename', 'page_index_to_copy', 'in_page_name', 'out_page_name'),
                         [
                             ('test1.vsdx', 0, None, 'Page-1-1'),
                             ('test1.vsdx', 0,  'newname', 'newname'),
                             ('test1.vsdx', 0, 'Page-1', 'Page-1-1'),
                             ('test1.vsdx', 1, None, 'Page-2-1'),
                             ('test1.vsdx', 1,  'newname', 'newname'),
                             ('test1.vsdx', 1, 'Page-1', 'Page-1-1'),
                         ])
def test_copy_page_naming(filename: str, page_index_to_copy:int, in_page_name: str, out_page_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_copy_page_naming.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index_to_copy]  # type: Page
        new_page = vis.copy_page(page, name=in_page_name)
        print(f"in_page_name:{in_page_name} out_page_name:{out_page_name} actual:{new_page.name}")
        assert new_page.name == out_page_name  # check new page has expected name
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(out_page_name)
        assert page  # check that page name persists through file save and open


@pytest.mark.parametrize(('filename', 'page_index_to_copy', 'page_position', 'out_page_index'),
                         [
                             ('test1.vsdx', 0, PagePosition.LAST, 3),
                             ('test1.vsdx', 0, PagePosition.BEFORE, 0),
                             ('test1.vsdx', 0, PagePosition.AFTER, 1),
                             ('test1.vsdx', 1, PagePosition.LAST, 3),
                             ('test1.vsdx', 1, PagePosition.BEFORE, 1),
                             ('test1.vsdx', 1, PagePosition.AFTER, 2),
                         ])
def test_copy_page_positions(filename: str, page_index_to_copy:int, page_position: PagePosition, out_page_index: int):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_copy_page_position.vsdx')
    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[page_index_to_copy]  # type: Page
        new_page = vis.copy_page(page, index=page_position)
        index = vis.pages.index(new_page)
        print(f"page_index_to_copy:{page_index_to_copy} page_position:{page_position} out_page_index:{out_page_index} actual:{index}")
        assert index == out_page_index  # check new page has expected index
        vis.save_vsdx(out_file)

    with VisioFile(out_file) as vis:
        page = vis.get_page_by_name(new_page.name)
        index = vis.pages.index(page)
        assert index == out_page_index  # check that page location persists through file save and open


# working with Shapes

@pytest.mark.parametrize(("filename", "shape_name"),
                         [("test1.vsdx", "Shape to copy"),
                          ("test4_connectors.vsdx", "Shape B")])
def test_vis_copy_shape(filename: str, shape_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_vis_copy_shape.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page = vis.pages[0]  # type: Page
        # find and copy shape by name
        s = page.find_shape_by_text(shape_name)  # type: Shape
        assert s  # check shape found
        print(f"Found shape id:{s.ID}")
        max_id = page.max_id

        # note = this does add the shape, but prefer Shape.copy() as per next test which wraps this and returns Shape
        page.set_max_ids()
        new_shape = vis.copy_shape(shape=s.xml, page=page)

        assert new_shape  # check copy_shape returns xml

        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.attrib['ID']}")
        assert int(new_shape.attrib.get('ID')) > int(s.ID)
        assert int(new_shape.attrib.get('ID')) > max_id

        new_shape_id = new_shape.attrib['ID']
        vis.save_vsdx(out_file)

    # re-open saved file and check it is changed as expected
    with VisioFile(out_file) as vis:
        page = vis.pages[0]
        s = page.find_shape_by_id(new_shape_id)
        assert s


@pytest.mark.parametrize(("filename", "shape_name"),
                         [("test1.vsdx", "Shape to copy"),
                          ("test2.vsdx", "Shape to copy")])
def test_copy_shape_other_page(filename: str, shape_name: str):
    out_file = os.path.join(basedir, 'out', f'{filename[:-5]}_test_copy_shape_other_page.vsdx')

    with VisioFile(os.path.join(basedir, filename)) as vis:
        page =  vis.pages[0]  # type: Page
        page2 = vis.pages[1]  # type: Page
        page3 = vis.pages[2]  # type: Page
        # find and copy shape by name
        s = page.find_shape_by_text(shape_name)  # type: Shape
        assert s  # check shape found
        shape_text = s.text
        print(f"Found shape id:{s.ID}")

        new_shape = vis.copy_shape(shape=s.xml, page=page2)
        assert new_shape  # check copy_shape returns xml
        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.attrib['ID']}")
        page2_new_shape_id = new_shape.attrib['ID']

        new_shape = vis.copy_shape(shape=s.xml, page=page3)
        assert new_shape  # check copy_shape returns xml
        print(f"created new shape {type(new_shape)} {new_shape} {new_shape.attrib['ID']}")
        page3_new_shape_id = new_shape.attrib['ID']

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
