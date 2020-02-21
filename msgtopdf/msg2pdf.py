import click
from colorama import init, Fore
from msgtopdf import Msgtopdf
from pathlib import Path, PurePath

# Initialise colorama
init()


@click.command()
@click.option(
    "-f", "--file", "path_type", flag_value="filename", help="Convert a file to pdf."
)
@click.option(
    "-d",
    "--directory",
    "path_type",
    flag_value="directory",
    help="Convert all msg files in directory to pdf.",
)
@click.argument("path", type=click.Path(exists=True, resolve_path=True))
def msg2pdf(path_type, path):
    if path_type == "filename":
        print("process file", path)
        convert_file(path)
    if path_type == "directory":
        print("process directory", path)


def convert_file(filename):
    try:
        print(filename)
        f = Msgtopdf(filename)
        f.email2pdf()
        print(Fore.GREEN + f"Converted {filename} to PDF!")
    except:
        print(Fore.RED + f"Filename is invalid, enter a valid filename!")


# @msg2pdf.command()
# @click.argument(
#     "directory",
#     # help="Name of directory to convert all msg files to pdf",
#     type=click.Path(exists=True, resolve_path=True),
# )
# def convert_folder(directory):
#     """Name of directory to convert all msg files to pdf"""
#     print(directory)


if __name__ == "__main__":
    msg2pdf()
