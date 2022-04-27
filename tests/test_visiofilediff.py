import pprint
import pytest
import os

from vsdx import VisioFile
from vsdx.vsdxdiff import VisioFileDiff


# code to get basedir of this test file in either linux/windows
basedir = os.path.dirname(os.path.relpath(__file__))


@pytest.mark.parametrize(("filename_a", "filename_b"),
                         [("test1.vsdx", "test2.vsdx" ),
                          ("test1.vsdx", "test4_connectors.vsdx" )])
def test_create_visiodiff(filename_a: str, filename_b: str):
    filepath_a = os.path.join(basedir, filename_a)
    filepath_b = os.path.join(basedir, filename_b)
    print(basedir)
    print(filepath_a)
    fd = VisioFileDiff(filepath_a, filepath_b)
    print(f"fd={fd}")
    print(f"Added in {filename_b} {fd.added_members()}")
    print(f"Removed in {filename_b} {fd.removed_members()}")
    for m in fd.common_members():
        print(f"\n\n{m}")
        pprint.pprint(fd.diffs.get(m))


# next test, open file, set text of shape, save as - then compare the two
@pytest.mark.skip
@pytest.mark.parametrize(("filename_a", "filename_b"),
                         [("test1.vsdx", "test1_outfile.vsdx" ),
                          ("test2.vsdx", "test2_outfile.vsdx" ),
                          ])
def test_visiodiff_before_after(filename_a: str, filename_b: str):
    filepath_a = os.path.join(basedir, filename_a)
    filepath_b = os.path.join(basedir, 'out', filename_b)
    with VisioFile(filepath_a) as vis:
        print(f"saving as {filepath_b}")
        vis.save_vsdx(filepath_b)

    fd = VisioFileDiff(filepath_a, filepath_b)

    print(f"fd={fd}")
    print(f"Added in {filename_b} {fd.added_members()}")
    print(f"Removed in {filename_b} {fd.removed_members()}")
    for m in fd.common_members():
        diff = fd.diffs.get(m)
        print(f"\n\n{m} {type(diff)} len:{len(diff) if diff else 0}")
        if diff:
            num = 0
            for l in diff:
                num += 1
                print(f"{num} {l}")


@pytest.mark.skip
@pytest.mark.parametrize(("filename_a", "filename_b"),
                         [
                             ("test4_connectors_out.vsdx", "test4_connectors_added.vsdx" ),
                          ])
def test_visiodiff_two_files(filename_a: str, filename_b: str):
    filepath_a = os.path.join(basedir, filename_a)
    filepath_b = os.path.join(basedir, filename_b)

    fd = VisioFileDiff(filepath_a, filepath_b)

    print(f"fd={fd}")
    print(f"Added in {filename_b} {fd.added_members()}")
    print(f"Removed in {filename_b} {fd.removed_members()}")
    for m in fd.common_members():
        diff = fd.diffs.get(m)
        print(f"\n\n{m} {type(diff)} len:{len(diff) if diff else 0}")
        if diff:
            num = 0
            for l in diff:
                num += 1
                print(f"{num} {l}")