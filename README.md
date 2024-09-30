# `certificates`

`certificates` is a purpose-built command-line application to generate certificates to be issued to interns completing their internship at Indian Institute of Information Technology Dharwad. Although it is specific to the needs of IIIT Dharwad, it can be easily adapted to other similar needs.

The general ideas on which it works are as follows:

1. The template for the certificates is a Microsoft Word `.docx` file (hereafter called the **template file**) and has all the static text, images (possibly including images of signatures) in the required character and paragraph formatting.
2. The template file contains **merge fields** repesentinf the keys which will be replaced by values from a Python dictionary with the matching key.
3. The template file is populated with data from a Python dictionary and a new Microsoft Word `.docx` file is created. It is then converted to a PDF file using LibreOffice, which must be installed.
4. A PDF file with a QR code containing all the fields is created and merged with the PDF file containing the certificate.
5. The certificate file with the QR code is then assigned a owner password to prevent editing and copying of the PDF file. However, no restriction is placed on printing the certificate.

# Dependencies
This project depends on the following Python packages:

1. Pandas: Data is input from a Microsoft Excel `.xlsx` file, with each row resulting in one certificate. Column names represent the names of merge fields in the template file.
2. openpyxl: It is a dependency of Pandas when you wish to read and write Microsoft Excel files.
3. docx-mailmerge2: It is used to merge data from a Python dictionary to populate the template file and generate one certificate.
4. pikepdf: It is used to set owner password to PDF files.
5. segno: It is used to generate QR codes in PNG format
6. PIL: It is used to convert QR code in PNG format to PDF.
7. toml: It is used to maintain a configuration file
8. click: It is used to define command line options and arguments.

In addition, this script requires LibreOffice (or OpenOffice) to be installed as it uses the application `soffice` to convert Microsoft Word `.docx` file to PDF file. The path to `soffice` is hardcoded in the script and must be changed appropriately. The script checks for the existence of `soffice` at the start and aborts if it is not found at the defined path.

# Installation
Follow these steps:

1. Create a separate directory for the script and within that directory, create a virtual environment using Python 3.12+
2. Install either `uv` or `pip-tools` using `pipx` if not already done.
3. Clone the Github repository
4. Create `requirements.txt` using `uv` or `pip-compile`. `uv pip compile requirements.in -o requirements.txt` or `pip-compile requirements.in -o requirements.txt`
5. Install required packages using `uv` or `pi-sync`. `uv pip sync requirements.txt` or `pip-sync requirements.txt`
6. Check if LibreOffice is installed, and if not install it from here: `https://www.libreoffice.org/`. Determine the path where `soffice` is installed. On GNU/Linux, use the `which soffice` command and on Windows, use `where soffice` command. If it is already on the `PATH` environment variable, it will be located. Else, use any method to determine the path to `soffice`, such as, looking up the properties of LibreOffice in your Windows Start menu.

# Input Data

# Executing the Script

