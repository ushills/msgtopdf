import os
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
    def __init__(self, msgfile):
        self.msgfile = msgfile
        self.directory = PurePath(self.msgfile).parent
        self.file = PurePath(self.msgfile).name
        self.save_path = self.__define_save_path()
        print(self.msgfile)
        print(self.directory)
        print(self.file)
        print(self.save_path)

    def raw_email_body(self):
        msg = outlook.OpenSharedItem(self.msgfile)
        if msg.BodyFormat == 2:
            body = msg.HTMLBody
        elif msg.BodyFormat == 3:
            body = msg.RTFBody
        else:
            body = msg.Body
        return raw_body

    def email_body(self):
        os.mkdir(self.save_path)
        print(f"Created folder: {save_path}")
        pass

    def __define_save_path(self):
        msgfile_name = self.file.split(".msg")[0]
        msgfile_folder = clean_path(msgfile_name)
        save_path = os.path.join(self.directory, msgfile_folder)
        return save_path


def extract_email_attachments(directory, msgfile):
    msg_path = os.path.join(directory, msgfile)
    msg = outlook.OpenSharedItem(msg_path)
    count_attachments = msg.Attachments.Count
    if count_attachments > 0:
        for item in range(count_attachments):
            filename = msg.Attachments.Item(item + 1).Filename
            save_path = create_save_path(directory, msgfile)
            msg.Attachments.Item(item + 1).SaveAsFile(save_path + "\\" + filename)


def convert_CID_image(input_file, output_file):
    with open(output_file, "w") as ofile:
        with open(input_file) as ifile:
            for line in ifile:
                if not line.rstrip():
                    continue
                else:
                    line = line.rstrip()
                    line = replace_CID(line) + "\n"
                    ofile.write(line)


def replace_CID(body):
    try:
        p = re.compile(r"cid:([^\"@]*)[^\"]*")
        m = p.search(body)
        r = p.sub((m.groups()[0]), body)
        return r
    except:
        return body


def save_email_body(directory, msgfile):
    input_file = "msgoutput.html"
    output_file = "output.html"
    msg_path = os.path.join(directory, msgfile)
    clean_body = extract_email_body(msg_path)
    save_path = create_save_path(directory, msgfile)
    input_file = os.path.join(save_path, input_file)
    outputfile = os.path.join(save_path, output_file)
    with open(outputfile, "w") as f:
        for char in clean_body:
            try:
                f.write(char)
            except:
                pass
    convert_CID_image(input_file, output_file)


def clean_path(path):
    c_path = re.sub(r'[\\/\:*"<>\|\.%\$\^&£]', "", path)
    c_path = re.sub(r"[ ]{2,}", "", c_path)
    c_path = c_path.strip()
    return c_path


def create_folder_structure(directory, msgfile):
    save_path = create_save_path(directory, msgfile)
    os.mkdir(save_path)
    print(f"Created folder: {save_path}")
    # if email_has_attachements(directory, msgfile):
    #     attachment_folder = os.path.join(save_path, "attachments")
    #     os.mkdir(attachment_folder)
    #     print("Created attachments folder")


def create_save_path(directory, msgfile):
    msgfile_name = msgfile.split(".msg")[0]
    msgfile_folder = clean_path(msgfile_name)
    save_path = os.path.join(directory, msgfile_folder)
    return save_path


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
