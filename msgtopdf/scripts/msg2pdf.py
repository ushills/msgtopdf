import click
from colorama import init, Fore
from msgtopdf.msgtopdf import Msgtopdf

from pathlib import Path, PurePath

# Initialise colorama
init()


@click.command()
@click.version_option()
@click.option(
    "-f",
    "--file",
    "path_type",
    flag_value="filename",
    help="Convert an individual file PATH to pdf.",
)
@click.option(
    "-d",
    "--directory",
    "path_type",
    flag_value="directory",
    help="Convert all msg files in directory PATH to pdf.",
)
@click.argument("path", type=click.Path(exists=True, resolve_path=True))
def cli(path_type, path):
    """msg2pdf converts Outlook email messages (msg) to pdf.\n
    The output is a folder for each email using the email subject as the folder name
    inculding a pdf of the email and all attachments.\n
    Inline images are included in the email pdf."""
    if path_type == "filename":
        convert_file(path)
    if path_type == "directory":
        convert_directory(path)


def convert_file(filename):
    try:
        f = Msgtopdf(filename)
        f.email2pdf()
        print(Fore.GREEN + f"Converted {filename} to PDF!" + Fore.RESET)
    except:
        print(Fore.RED + f"Something went wrong!" + Fore.RESET)


def convert_directory(directory):
    msg_files = list(Path(directory).glob("**/*.msg"))
    for f in msg_files:
        convert_file(f)

