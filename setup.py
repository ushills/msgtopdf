import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="msgtopdf_pkg_ushills",
    version="0.0.1",
    author="Ian Hill",
    author_email="web@ushills.co.uk",
    description="Convert Outlook msg to PDF",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/ushills/msgtopdf",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "Development Status :: 4 - Beta",
        "License :: OSI Approved :: MIT License",
        "Environment :: Win32 (MS Windows)",
        "Operating System :: Microsoft :: Windows",
    ],
    python_requires=">=3.6",
)
