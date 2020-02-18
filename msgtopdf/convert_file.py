import click
from colorama import init, Fore
from msgtopdf import Msgtopdf
from pathlib import Path, PurePath

init()


@click.command()
@click.option(
    "-f",
    "--file",
    "filename",
    help="Name of the file to convert",
    type=click.Path(exists=True, resolve_path=True),
)
def convert_file(filename):
    try:
        print(filename)
        f = Msgtopdf(filename)
        f.email2pdf()
        print(Fore.GREEN + f"Converted {filename} to PDF!")
    except:
        print(Fore.RED + f"Filename is invalid, enter a valid filename!")


if __name__ == "__main__":
    convert_file()
