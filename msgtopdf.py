import cgi
import os
import re
import shutil
from pathlib import Path, PurePath, WindowsPath

import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


class MsgtoPdf:
    def __init__(self, msgfile):
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

        temp_outputfile = PurePath(self.save_path, "~temp")
        with open(temp_outputfile, "w", encoding="utf-8") as f:
            f.write(full_email_body)

        # clean the file to correct the CID image tags to normal image files
        clean_temp_outputfile = PurePath(self.save_path, "~cleantemp")
        self.convert_CID_image(temp_outputfile, clean_temp_outputfile)
        os.remove(temp_outputfile)

        # rename the output file with the correct extension
        # email_body_file = PurePath(
        #     self.save_path, self.file_name + "." + self.email_format
        # )
        # print(email_body_file)
        # os.rename(clean_temp_outputfile, email_body_file)
        print(self.image_files)

    def __define_save_path(self):
        msgfile_name = self.file.split(".msg")[0]
        msgfile_folder = clean_path(msgfile_name)
        save_path = PurePath(self.directory, msgfile_folder)
        return save_path

    def add_header_information(self):
        html_str = """
        <p style="font-family: Arial;font-size: 12.0pt">
        <strong>From:</strong>               {sender}</br>
        <strong>Sent:</strong>               {sent}</br>
        <strong>To:</strong>                 {to}</br>
        <strong>Cc:</strong>                 {cc}</br>
        <strong>Subject:</strong>            {subject}</p>
        <hr>
        """
        formatted_html = html_str.format(
            sender=self.msg.SenderName,
            sent=self.msg.SentOn,
            to=self.msg.To,
            cc=self.msg.CC,
            subject=self.msg.Subject,
            attachments=self.msg.Attachments,
        )
        return formatted_html

    def convert_CID_image(self, input_file, output_file):
        with open(output_file, "a") as ofile:
            with open(input_file) as ifile:
                for line in ifile:
                    if not line.rstrip():
                        continue
                    else:
                        line = line.rstrip()
                        line = self.replace_CID(line) + "\n"
                        ofile.write(line)

    def replace_CID(self, line):
        self.image_files = []
        try:
            p = re.compile(r"cid:([^\"@]*)[^\"]*")
            m = p.search(line)
            self.image_files.append(m.groups()[0])
            line = p.sub((m.groups()[0]), line)
            return line
        except:
            return line


def extract_email_attachments(directory, msgfile):
    msg_path = os.path.join(directory, msgfile)
    msg = outlook.OpenSharedItem(msg_path)
    count_attachments = msg.Attachments.Count
    if count_attachments > 0:
        for item in range(count_attachments):
            filename = msg.Attachments.Item(item + 1).Filename
            save_path = create_save_path(directory, msgfile)
            msg.Attachments.Item(item + 1).SaveAsFile(save_path + "\\" + filename)


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
