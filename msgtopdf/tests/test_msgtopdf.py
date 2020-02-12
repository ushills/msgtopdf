import pytest
import pathlib
from unittest.mock import MagicMock
import win32com.client


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

    def test_replace_CID_single_CID(self):
        email = MsgtoPdf("C:/test/email.msg")
        line = '<img src="cid:image001.png@01D54589.B5E9EB60">'
        assert email.replace_CID(line) == '<img src="image001.png">'

    def test_replace_CID_alt_CID(self):
        email = MsgtoPdf("C:/test/email.msg")
        line = '<img width="1135" height="571" style="width:11.8229in;height:5.9479in" id="Picture_x0020_1" src="cid:image004.png@01D543A2.096B0830" alt="cid:image004.png@01D543A2.096B0830">'
        assert (
            email.replace_CID(line)
            == '<img width="1135" height="571" style="width:11.8229in;height:5.9479in" id="Picture_x0020_1" src="image004.png" alt="image004.png">'
        )

    def test_replace_CID_multiple_CID(self):
        email = MsgtoPdf("C:/test/email.msg")
        line = '<img src="cid:image001.png@01D54589.B5E9EB60"><img src="cid:image002.png@01D54589.B5E9EB60">'
        assert (
            email.replace_CID(line)
            == '<img src="image001.png"><img src="image002.png">'
        )
        assert email.image_files == ["image001.png", "image002.png"]

    def test_replace_CID_no_replace(self):
        email = MsgtoPdf("C:/test/email.msg")
        body = "<p>Not an image</p>"
        assert email.replace_CID(body) == "<p>Not an image</p>"

    def test_clean_path(self):
        email = MsgtoPdf("C:/test/email.msg")
        path = r"RE:/ test dirty path ^"
        assert email.clean_path(path) == "RE test dirty path"

