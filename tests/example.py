from msgtopdf import Msgtopdf
from pathlib import Path


def main():
    directory = Path.cwd()
    msgfile = "file.msg"
    msgfile = Path(directory, msgfile)
    email = Msgtopdf(msgfile)
    email.email2pdf()


if __name__ == "__main__":
    main()
