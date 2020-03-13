import re
from pathlib import Path, PurePath
from subprocess import run

import win32com.client


class Msgtopdf:
    def __init__(self, msgfile):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.msgfile = PurePath(msgfile)
        self.directory = PurePath(self.msgfile).parent
        self.file = PurePath(self.msgfile).name
        self.file_name = self.file.split(".msg")[0]
        self.save_path = self.__define_save_path()
        self.msg = outlook.OpenSharedItem(self.msgfile)
        self.raw_body = None

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

    def email2pdf(self):
        Path.mkdir(Path(self.save_path))
        html_header = self.__add_header_information()
        raw_email_body = self.raw_email_body
        full_email_body = html_header + raw_email_body
        clean_email_body = self.replace_CID(full_email_body)
        self.html_body_file = PurePath(self.save_path, self.file_name + ".html")
        self.extract_email_attachments()
        # convert_html_to_pdf(clean_email_body, self.html_body_file)
        with open(self.html_body_file, "w", encoding="utf-8") as f:
            f.write(clean_email_body)
        # save pdf copy using wkhtmltopdf
        run(
            [
                "wkhtmltopdf",
                "--log-level",
                "warn",
                "--encoding",
                "utf-8",
                "--footer-font-size",
                "6",
                "--footer-line",
                "--footer-center",
                "[page] / [topage]",
                self.html_body_file,
                PurePath(self.save_path, self.file_name + ".pdf"),
            ]
        )
        self.__delete_redundant_files()

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
        msgfile_folder = self.clean_path(msgfile_name)
        save_path = Path(self.directory, msgfile_folder)
        increment = 0
        while save_path.exists() and save_path.is_dir():
            save_path = Path(self.directory, "%s (%d)" % (msgfile_folder, increment))
            increment += 1
        return PurePath(save_path)

    def __add_header_information(self):
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
        r = p.sub(self.__return_image_reference, body)
        return r

    def __return_image_reference(self, match):
        value = str(match.groups()[0])
        if value not in self.image_files:
            self.image_files.append(value)
        return value

    def __delete_redundant_files(self):
        Path.unlink(Path(self.html_body_file))
        for f in self.image_files:
            image_full_path = Path(self.save_path, f)
            if Path.exists(image_full_path):
                Path.unlink(image_full_path)

    def clean_path(self, path):
        c_path = re.sub(r'[\\/\:*"<>\|\.%\$\^&£]', "", path)
        c_path = re.sub(r"[ ]{2,}", "", c_path)
        c_path = c_path.strip()
        return c_path

