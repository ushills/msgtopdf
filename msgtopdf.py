import os
import re
import shutil
from pathlib import Path, PurePath, WindowsPath
from subprocess import run

import win32com.client


class MsgtoPdf:
    def __init__(self, msgfile):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.msgfile = msgfile
        self.directory = PurePath(self.msgfile).parent
        self.file = PurePath(self.msgfile).name
        self.file_name = self.file.split(".msg")[0]
        self.save_path = self.__define_save_path()
        self.msg = outlook.OpenSharedItem(self.msgfile)

    def raw_email_body(self):
        if self.msg.BodyFormat == 2:
            body = self.msg.HTMLBody
            self.email_format = "html"
        elif self.msg.BodyFormat == 3:
            body = self.msg.RTFBody
            self.email_format = "html"
        else:
            body = self.msg.Body
            self.email_format = "txt"
        self.raw_body = body
        return self.raw_body

    def save_email_body(self):
        os.mkdir(self.save_path)
        print(f"Created folder: {self.save_path}")
        html_header = self.add_header_information()
        raw_email_body = self.raw_email_body()
        full_email_body = html_header + raw_email_body
        clean_email_body = self.replace_CID(full_email_body)
        body_file = PurePath(self.save_path, self.file_name + ".html")
        self.extract_email_attachments()
        # convert_html_to_pdf(clean_email_body, body_file)
        with open(body_file, "w", encoding="utf-8") as f:
            f.write(clean_email_body)
        # save pdf copy using wkhtmltopdf
        run(
            [
                "wkhtmltopdf",
                "--encoding",
                "utf-8",
                "--footer-font-size",
                "6",
                "--footer-line",
                "--footer-center",
                "[page] / [topage]",
                body_file,
                PurePath(self.save_path, self.file_name + ".pdf"),
            ]
        )

    def extract_email_attachments(self):
        count_attachments = self.msg.Attachments.Count
        if count_attachments > 0:
            for item in range(count_attachments):
                attachment_filename = self.msg.Attachments.Item(item + 1).Filename
                self.msg.Attachments.Item(item + 1).SaveAsFile(
                    PurePath(self.save_path, attachment_filename)
                )

    def __define_save_path(self):
        msgfile_name = self.file.split(".msg")[0]
        msgfile_folder = clean_path(msgfile_name)
        save_path = PurePath(self.directory, msgfile_folder)
        return save_path

    def add_header_information(self):
        html_str = """
        <head>
        <meta charset="UTF-8">
        <base href="{base_href}">
        <p style="font-family: Arial;font-size: 11.0pt">
        </head>
        <strong>From:</strong>               {sender}</br>
        <strong>Sent:</strong>               {sent}</br>
        <strong>To:</strong>                 {to}</br>
        <strong>Cc:</strong>                 {cc}</br>
        <strong>Subject:</strong>            {subject}</p>
        <hr>
        """
        formatted_html = html_str.format(
            base_href="file:///" + str(self.save_path) + "\\",
            sender=self.msg.SenderName,
            sent=self.msg.SentOn,
            to=self.msg.To,
            cc=self.msg.CC,
            subject=self.msg.Subject,
            attachments=self.msg.Attachments,
        )
        return formatted_html

    def replace_CID(self, body):
        self.image_files = []
        # search for cid:(capture_group)@* upto "
        p = re.compile(r"cid:([^\"@]*)[^\"]*")
        r = p.sub(self.return_image_reference, body)
        # self.image_files.append(m.groups()[0])
        # print(r)
        print(self.image_files)
        return r

    def return_image_reference(self, match):
        value = str(match.groups()[0])
        if value not in self.image_files:
            self.image_files.append(value)
        return value


def clean_path(path):
    c_path = re.sub(r'[\\/\:*"<>\|\.%\$\^&£]', "", path)
    c_path = re.sub(r"[ ]{2,}", "", c_path)
    c_path = c_path.strip()
    return c_path


def email_has_attachements(directory, msgfile):
    msg_path = os.path.join(directory, msgfile)
    msg = outlook.OpenSharedItem(msg_path)
    if msg.Attachments.Count > 0:
        return True


def process_email(directory, msgfile):
    create_folder_structure(directory, msgfile)
    save_email_body(directory, msgfile)
    extract_email_attachments(directory, msgfile)


def main():
    directory = os.getcwd()
    msgfile = "file.msg"
    msgfile = os.path.join(directory, msgfile)
    email = MsgtoPdf(msgfile)
    email.save_email_body()


if __name__ == "__main__":
    main()
