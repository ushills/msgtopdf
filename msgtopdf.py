import os
import shutil
import win32com.client
import re
from pathlib import Path, WindowsPath, PurePath

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# print(msg.SenderName)
# print(msg.SenderEmailAddress)
# print(msg.SentOn)
# print(msg.To)
# print(msg.CC)
# print(msg.BCC)
# print(msg.Subject)
# BodyFormat 2 = HTML, 1 = Plain, 3 = RTF
# if msg.BodyFormat == 2:
#     print(msg.HTMLBody)
#     html = msg.HTMLBody
#     with open("msgbody.html", "w") as f:
#         f.write(html)
#     # with open("msgoutput.html", "w") as f:
#     #     for line in html:
#     #         try:
#     #             f.write(line)
#     #         except:
#     #             pass
# elif msg.BodyFormat == 3:
#     print(msg.RTFBody)
# else:
#     print(msg.Body)

# # Attachments
# # count_attachments = msg.Attachments.Count
# # if count_attachments > 0:
# #     for item in range(count_attachments):
# #         filename = msg.Attachments.Item(item + 1).Filename
# #         print(filename)
# #         msg.Attachments.Item(item + 1).SaveAsFile(os.getcwd() + "\\" + filename)

# del outlook, msg


class MsgtoPdf:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

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
        raw_email_body = self.raw_email_body()
        temp_outputfile = PurePath(self.save_path, "~temp")
        with open(temp_outputfile, "w") as f:
            for char in raw_email_body:
                try:
                    f.write(char)
                except:
                    pass

        # clean the file to correct the CID image tags to normal image files
        clean_temp_outputfile = PurePath(self.save_path, "~cleantemp")
        html_str = self.add_header_information()
        with open(clean_temp_outputfile, "w") as ofile:
            ofile.write(html_str)
        convert_CID_image(temp_outputfile, clean_temp_outputfile)
        os.remove(temp_outputfile)

        # rename the output file with the correct extension
        email_body_file = PurePath(
            self.save_path, self.file_name + "." + self.email_format
        )
        print(email_body_file)
        os.rename(clean_temp_outputfile, email_body_file)

    def __define_save_path(self):
        msgfile_name = self.file.split(".msg")[0]
        msgfile_folder = clean_path(msgfile_name)
        save_path = PurePath(self.directory, msgfile_folder)
        return save_path

    def add_header_information(self):
        html_str = """
        <p>From:               {sender}</p>
        <p>Sent:               {sent}</p>
        <p>To:                 {to}</p>
        <p>Cc:                 {cc}</p>
        <p>Subject:            {subject}</p>
        <p>Attachments:        {attachments}<p>
        </br>
        </br>
        """
        formatted_html = html_str.format(
            sender=self.msg.SenderName,
            sent=self.msg.SentOn,
            to=self.msg.To,
            cc=self.msg.CC,
            subject=self.msg.Subject,
            attachments=self.msg.Attachments,
        )
        print(formatted_html)
        return formatted_html


def convert_CID_image(input_file, output_file):
    with open(output_file, "a") as ofile:
        with open(input_file) as ifile:
            for line in ifile:
                if not line.rstrip():
                    continue
                else:
                    line = line.rstrip()
                    line = replace_CID(line) + "\n"
                    ofile.write(line)


def replace_CID(line):
    try:
        p = re.compile(r"cid:([^\"@]*)[^\"]*")
        m = p.search(line)
        r = p.sub((m.groups()[0]), line)
        return r
    except:
        return line


def add_header_information():
    pass


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
    # raw_body = email.raw_email_body()
    # print(raw_body)


if __name__ == "__main__":
    main()
