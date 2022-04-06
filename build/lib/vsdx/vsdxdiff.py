import difflib
import zipfile
import shutil
import os


class VisioFileDiff:
    """Compares two vsdx files

        :param filepath_a: file path of the first :class:`VisioFile` was created from
        :type filepath_a: str
        :param filepath_b: file path of the second :class:`VisioFile` was created from
        :type filepath_b: str
        """
    def __init__(self, filepath_a: str, filepath_b: str):
        if filepath_a == filepath_b:
            raise ValueError('The two file paths should be different')
        if not filepath_a.lower().endswith('.vsdx') or not filepath_b.lower().endswith('.vsdx'):
            raise ValueError('Both files should be vsdx files')

        # load contents of each file
        self.filepath_a = filepath_a
        self.contents_a = self.extract_file_data(filepath_a)
        self.filepath_b = filepath_b
        self.contents_b = self.extract_file_data(filepath_b)

        # check if each file has same contents
        self.diffs = self.get_file_diffs()

    def __str__(self):
        return f"VisioFileDiff(a={self.filepath_a}, b={self.filepath_b})"

    def get_file_diffs(self):
        common_members = self.common_members()
        diffs = {}
        d = difflib.Differ()
        for member_name in common_members:
            data_a = self.contents_a.get(member_name)
            data_b = self.contents_b.get(member_name)
            if data_a and data_b and data_a != data_b:  # only add diff if contents are not the same
                diffs[member_name] = list(d.compare(data_a, data_b))
        return diffs

    def common_members(self) -> list:
        # return a sorted list of members (file paths)
        common_members = list(set(self.contents_a.keys()).union(set(self.contents_b.keys())))
        common_members.sort()
        return common_members

    def compare_members(self) -> bool:
        # return True if same, False if different
        if self.contents_a.keys() == self.contents_b.keys():
            return True
        else:
            return False

    def added_members(self) -> set:
        # list members in file b that are not in file a
        members_a = set(self.contents_a.keys())
        members_b = set(self.contents_b.keys())
        return members_b.difference(members_a)

    def removed_members(self) -> set:
        # list members in file b that are not in file a
        members_a = set(self.contents_a.keys())
        members_b = set(self.contents_b.keys())
        return members_a-members_b

    @staticmethod
    def extract_file_data(file_path: str) -> dict:
        # open a vsdx file (or other zip based format) and return a dictionary of file contents by file_path
        directory = os.path.abspath(file_path)[:-5]  # -5 to remove '.vsdx' from filename
        with zipfile.ZipFile(file_path, "r") as zip_ref:
            extracted_file_paths = zip_ref.namelist()
            zip_ref.extractall(directory)

        # process data in directory
        file_contents = {}
        for extracted_file_path in extracted_file_paths:
            #print(f"Opening {os.path.join(directory, extracted_file_path)}")
            full_path = os.path.join(directory, extracted_file_path)
            try:
                if not os.path.isdir(full_path):
                    with open(full_path, mode='r') as f:
                        file_data = f.readlines()
                    file_contents[extracted_file_path] = file_data
                    #print(f"Opened and read contents of {extracted_file_path}")
            except UnicodeDecodeError as e:
                file_contents[extracted_file_path] = "Unable to decode file."
                print(f"Failed to read file: {full_path}")
            except PermissionError as e:
                file_contents[extracted_file_path] = "Unable to open file."
                print(f"Failed to open file (PermissionError): {full_path}")
        try:
            # Remove extracted folder
            shutil.rmtree(directory)
        except (FileNotFoundError, PermissionError) as e:
            print(f"Error shutil.rmtree({directory}) {e}")

        return file_contents