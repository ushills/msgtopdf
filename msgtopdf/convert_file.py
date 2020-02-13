import click
from colorama import init, Fore
from msgtopdf import Msgtopdf
from pathlib import Path, PurePath

init()


@click.command()
@click.option("--file", help="Name of file to convert")
def convert_file(file):
    try:
        msgfile = Path(file)
        msgfile = msgfile.absolute()
        f = Msgtopdf(msgfile)
        f.email2pdf()
        print(Fore.GREEN + f"Converted {file} to PDF!")
    except:
        print(Fore.RED + f"Filename is invalid, enter a valid filename!")


if __name__ == "__main__":
    convert_file()
