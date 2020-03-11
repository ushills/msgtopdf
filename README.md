![Test msgtopdf](https://github.com/ushills/msgtopdf/workflows/Test%20msgtopdf/badge.svg) [![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

# Converts Outlook .msg files to PDF

`msgtopdf` is a Python 3 module to convert Outlook `.msg` files to PDF and extract the attachments.  Unlike the majority of current modules `msgtopdf` maintains the formatting of HTML and RTF messages and embeds any inline images in the PDF output.

As the module uses the `win32com` library the host machine must have Outlook installed.

`msgtopdf` uses the [wkhtmltopdf](https://wkhtmltopdf.org/) tool to convert the HTML message to PDF and [wkhtmltopdf](https://wkhtmltopdf.org/) must be installed separately.

Currently `msgtopdf` extracts the message body and attachments to a new subfolder using the subject of the email as the folder name.

# Usage

## Module Usage

Example module usage is provided in the `tests/example.py` file.


## Command Line Usage

The command-line option `msg2pdf` will convert individual files or all `*.msg` files in a directory.

`msg2pdf --help` for options.

    Usage: msg2pdf [OPTIONS] PATH

    msg2pdf converts Outlook email messages (msg) to pdf.

    The output is a folder for each email using the email subject as the
    folder name including a pdf of the email and all attachments.

    Inline images are included in the email pdf.

    Options:
        -f, --file       Convert an individual file PATH to pdf.
        -d, --directory  Convert all msg files in directory PATH to pdf.
        --help           Show this message and exit.



# Requirements

Install the Windows binary release of [wkhtmltopdf](https://wkhtmltopdf.org/downloads.html)

Ensure that `wkhtmltopdf` command is found in your `PATH`.

This can be tested by entering `wkhtmltopdf --version` in your Command Prompt.

You should receive an output similar to the attached.


    Microsoft Windows [Version 6.1.7601]
    Copyright (c) 2009 Microsoft Corporation.  All rights reserved.

    C:\>wkhtmltopdf --version
    wkhtmltopdf 0.12.5 (with patched qt)

    C:\>
