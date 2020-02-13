from msgtopdf import MsgtoPdf
from pathlib import Path


def main():
    directory = Path.cwd()
    msgfile = "file.msg"
    msgfile = Path(directory, msgfile)
    email = MsgtoPdf(msgfile)
    email.save_email_body()


if __name__ == "__main__":
    main()
