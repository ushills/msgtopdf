from setuptools import setup, find_packages


with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="msgtopdf",
    version="0.1.4",
    author="Ian Hill",
    author_email="web@ushills.co.uk",
    description="Convert Outlook msg to PDF",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/ushills/msgtopdf",
    packages=find_packages(exclude=["tests"]),
    install_requires=["pywin32", "Click", "Colorama"],
    entry_points="""
        [console_scripts]
        msg2pdf=msgtopdf.scripts.msg2pdf:cli
    """,
    include_package_data=True,
    classifiers=[
        "Programming Language :: Python :: 3",
        "Development Status :: 4 - Beta",
        "License :: OSI Approved :: MIT License",
        "Environment :: Win32 (MS Windows)",
        "Operating System :: Microsoft :: Windows",
    ],
    python_requires=">=3.6",
)
