import pytest
import pathlib
from unittest.mock import MagicMock
import win32com.client

from msgtopdf.msgtopdf import Msgtopdf

mock_outlook = MagicMock()
win32com.client = mock_outlook


class Test_Msgtopdf:
    def test_init_directory(self):
        email = Msgtopdf("C:/test/email.msg")
        assert email.directory == pathlib.PurePath("C:/test")

    def test_init_file(self):
        email = Msgtopdf("C:/test/email.msg")
        assert email.file == "email.msg"

    def test_init_save_path(self):
        email = Msgtopdf("C:/test/email.msg")
        assert email.save_path == pathlib.PurePath("C:/test/email")

        p = pathlib.Path("C:/test/email/email")
        p.mkdir(exist_ok=True, parents=True)
        email = Msgtopdf("C:/test/email.msg")
        assert email.save_path == pathlib.PurePath("C:/test/email (0)")

        p1 = pathlib.Path("C:/test/email (0)")
        p1.mkdir(exist_ok=True, parents=True)
        email = Msgtopdf("C:/test/email.msg")
        assert email.save_path == pathlib.PurePath("C:/test/email (1)")

        p.rmdir()
        p1.rmdir()

    def test_replace_CID_single_CID(self):
        email = Msgtopdf("C:/test/email.msg")
        line = '<img src="cid:image001.png@01D54589.B5E9EB60">'
        assert email.replace_CID(line) == '<img src="image001.png">'

    def test_replace_CID_alt_CID(self):
        email = Msgtopdf("C:/test/email.msg")
        line = '<img width="1135" height="571" style="width:11.8229in;height:5.9479in" id="Picture_x0020_1" src="cid:image004.png@01D543A2.096B0830" alt="cid:image004.png@01D543A2.096B0830">'
        assert (
            email.replace_CID(line)
            == '<img width="1135" height="571" style="width:11.8229in;height:5.9479in" id="Picture_x0020_1" src="image004.png" alt="image004.png">'
        )

    def test_replace_CID_multiple_CID(self):
        email = Msgtopdf("C:/test/email.msg")
        line = '<img src="cid:image001.png@01D54589.B5E9EB60"><img src="cid:image002.png@01D54589.B5E9EB60">'
        assert (
            email.replace_CID(line)
            == '<img src="image001.png"><img src="image002.png">'
        )
        assert email.image_files == ["image001.png", "image002.png"]

    def test_replace_CID_no_replace(self):
        email = Msgtopdf("C:/test/email.msg")
        body = "<p>Not an image</p>"
        assert email.replace_CID(body) == "<p>Not an image</p>"

    def test_clean_path(self):
        email = Msgtopdf("C:/test/email.msg")
        path = r"RE:/ test dirty path ^"
        assert email.clean_path(path) == "RE test dirty path"

    def test_clean_path(self):
        email = Msgtopdf("C:/test/email.msg")
        path = r"RE:/ test dirty path ^"
        assert email.clean_path(path) == "RE test dirty path"

    def test___delete_redundant_files(self):
        email = Msgtopdf("C:/test/email.msg")
        email.save_path = pathlib.PurePath("./tests/")
        email.image_files = ["exists.png", "does_not_exist.png"]
        email.html_body_file = "./tests/html_body.html"
        # create temporary files for deletion
        open("./tests/html_body.html", "w+")
        open("./tests/exists.png", "w+")
        assert email.image_files == ["exists.png", "does_not_exist.png"]
        email._Msgtopdf__delete_redundant_files()

    def test_raw_email_body_html(self):
        email = Msgtopdf("C:/test/email.msg")
        email.msg.BodyFormat = 2
        email.raw_email_body()
        assert email.email_format == "html"
        email.msg.BodyFormat = 3
        email.raw_email_body()
        assert email.email_format == "html"
        email.msg.BodyFormat = 1
        email.raw_email_body()
        assert email.email_format == "txt"

