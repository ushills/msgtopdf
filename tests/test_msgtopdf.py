import pytest
import pathlib
from unittest.mock import MagicMock
import win32com.client


from msgtopdf import clean_path
from msgtopdf import MsgtoPdf

mock_outlook = MagicMock()
win32com.client = mock_outlook


class Test_MsgtoPdf:
    def test_init_directory(self):
        email = MsgtoPdf("C:/test/email.msg")
        assert email.directory == pathlib.PurePath("C:/test")

    def test_init_file(self):
        email = MsgtoPdf("C:/test/email.msg")
        assert email.file == "email.msg"

    def test_init_save_path(self):
        email = MsgtoPdf("C:/test/email.msg")
        assert email.save_path == pathlib.PurePath("C:/test/email")

    # def test_replace_CID(self):
    #     line = 'src="cid:image002.png@01D54589.B5E9EB60"'
    #     assert email.replace_CID(line) == 'src="image002.png"'

    #     line = '<img width="1135" height="571" style="width:11.8229in;height:5.9479in" id="Picture_x0020_1" src="cid:image004.png@01D543A2.096B0830" alt="cid:image004.png@01D543A2.096B0830">'
    #     assert (
    #         self.replace_CID(line)
    #         == '<img width="1135" height="571" style="width:11.8229in;height:5.9479in" id="Picture_x0020_1" src="image004.png" alt="image004.png">'
    #     )

    # def test_replace_CID_no_replace(self):
    #     line = "<p>Not an image</p>"
    #     assert email.replace_CID(line) == "<p>Not an image</p>"


def test_clean_path():
    path = r"RE:/ test dirty path ^"
    assert clean_path(path) == "RE test dirty path"

