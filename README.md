# Converts Outlook .msg files to PDF

`msgtopdf` is a Python 3 module to convert Outlook `.msg` files to PDF and extract the attachments.  Unlike the majority of current modules `msgtopdf` maintains the formatting of HTML and RTF messages and embeds any inline images in the PDF output.

As the module uses the `win32com` library the host machine must have Outlook installed.

`msgtopdf` uses the [wkhtmltopdf](https://wkhtmltopdf.org/) tool to convert the HTML message to PDF and [wkhtmltopdf](https://wkhtmltopdf.org/) must be installed separately.

Currently `msgtopdf` extracts the message body and attachments to a new subfolder using subject of the email as the folder name.

# Usage

Example usage is provided in the `tests/example.py` file.  The next plan on the timeline is to create a command line tool in which the user provides either a filename `-f` or directory `-d` and `msgtopdf` will process either a single `msg` file or all `msg` files in a directory.

# Requirements

Install the Windows binary release of [wkhtmltopdf](https://wkhtmltopdf.org/downloads.html)

Ensure that `wkhtmltopdf` command is found in your `PATH`.

This can be tested by entering `wkhtmltopdf --version` in your Command Prompt.

You should receive an output similar to the attached.

```
Microsoft Windows [Version 6.1.7601]
Copyright (c) 2009 Microsoft Corporation.  All rights reserved.

C:\>wkhtmltopdf --version
wkhtmltopdf 0.12.5 (with patched qt)

C:\>
```