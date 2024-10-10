# `certificates`

`certificates` is a purpose-built command-line interface (CLI) to generate certificates to be issued to interns completing their internship at Indian Institute of Information Technology Dharwad. Although it is specific to the needs of IIIT Dharwad, it can be easily adapted to other similar needs.

The general ideas on which it works are as follows:

1. The template for the certificates is a Microsoft Word `.docx` file (hereafter called the **template file**) and has all the static text, images (possibly including images of signatures) in the required character and paragraph formatting.
2. The template file contains [**merge fields**](https://support.microsoft.com/en-us/office/field-codes-mergefield-field-7a6d24a1-68a6-4b05-8359-1dc087daf4e6) repesenting the keys which will be replaced by values from a Python dictionary with the matching key.
3. The template file is populated with data from a Python dictionary and a new Microsoft Word `.docx` file is created. It is then converted to a PDF file using LibreOffice, which must be installed.
4. A PDF file with a QR code containing all the fields is created and merged with the PDF file containing the certificate.
5. The certificate file with the QR code is then assigned a owner password to prevent editing and copying of the PDF file. However, no restriction is placed on printing the certificate.

# Requirements
You will need the following:
1. [Python 3.12+](https://www.python.org) 
2. [Git](https://git-scm.com/) to clone the Github repo
3. One of [`uv`](https://git-scm.com/) or [`pip-tools`](https://github.com/jazzband/pip-tools/) to manage installation of required packages. You can manage it with just `pip` if you know Python well.
4. [LibreOffice](https://www.libreoffice.org/) or [OpenOffice](https://www.openoffice.org/) because this script uses the application `soffice` to convert Microsoft Word `.docx` file to PDF file. The path to `soffice` is hardcoded in the script and must be changed appropriately. The script checks for the existence of `soffice` at the start and aborts if it is not found at the defined path.

If you wish to use `uv` or `pip-tools`, it is best to install them using `pipx`.

# Dependencies
This project depends on the following Python packages:

1. [`pandas`](https://pandas.pydata.org/): Data is input from a Microsoft Excel `.xlsx` file, with each row resulting in one certificate. Column names represent the names of merge fields in the template file.
2. `openpyxl`: It is a dependency of Pandas when you wish to read and write Microsoft Excel files.
3. [`docx-mailmerge2`](https://github.com/iulica/docx-mailmerge): It is used to merge data from a Python dictionary to populate the template file and generate one certificate.
4. [`pikepdf`](https://github.com/pikepdf/pikepdf): It is used to set owner password to PDF files.
5. [`segno`](https://github.com/heuer/segno/): It is used to generate QR codes in PNG format
6. [`PIL`](https://python-pillow.org/): It is used to convert QR code in PNG format to PDF.
7. [`toml`](https://github.com/uiri/toml): It is used to maintain a configuration file
8. [`click`](https://click.palletsprojects.com/): It is used to define command line options and arguments.

# Installation
Follow these steps:

1. Create a separate directory for the script and within that directory, create a virtual environment using Python 3.12+ and activate it.
2. Use `uv` or `pip-tools` to manage package installation and updation. Install one of them `uv` or `pip-tools`. The best way is to use `pipx` to install them in their separate virtual environments and make them available to all projects.
3. Clone the Github repository `git clone https://github.com/satish-annigeri/certificates.git`.
4. Create `requirements.txt` using `uv` or `pip-compile` (`uv pip compile requirements.in -o requirements.txt` or `pip-compile requirements.in -o requirements.txt`)
5. Install required packages using `uv` or `pi-sync`. (`uv pip sync requirements.txt` or `pip-sync requirements.txt`)
6. Check if LibreOffice is installed, and if not, install it from here: `https://www.libreoffice.org/`. Determine the path where `soffice` is installed. On GNU/Linux, use the `which soffice` command and on Windows, use `where soffice` command. If it is already on the `PATH` environment variable, it will be located. Else, use any method to determine the path to `soffice`, such as, looking up the properties of LibreOffice in your Windows Start menu. Open the script `gencert.py` in your IDE and search for `libre_office_path` and change it appropriately.

# Input Data
Data is input in a Microsoft Excel `.xlsx` file. The following column names are mandatory:
1. `student_name`: Name of the student. If the student's name starts with one of `Mr.`, `Ms.`, or `Mrs.`, that portion, along with a trailing space if present, will be removed when used in printing the name in the certificate.
2. `gender`: Gender of the student. Can be `M` for male, `F` for female or any other letter or empty. This will determine whether the student's name in the certificate is prefixed with `Mr. `, `Ms. ` or left unchanged.
3. `start_date`: Start date of the internship.
4. `end_date`: Last date of the inetrnship.
5. `supervisor_name`: Name of the supervisor. Will be printed as is. Include `Dr. ` or `Prof. ` as the case may be when writing the name.
6. `designation`: Designation of the supervisor.
7. `department`: Department of the supervisor.
8. `project_title`: Title of the project on which the intern worked.

If any other columns are present, they will be read but not used.

# Executing the Script
Place the input data file in the same directory as the script and execute the command:

`(venv)>python gencert.py --no-final int_cert_data.xlsx`

where `int_cert_data.xlsx` is the Microsoft Excel file containing the data as described above. With the `--no-final` switch, only a dry run is executed, certificates will not be generated. The screen output gives you an idea if the data was arranged correctly in the input data file.

To generate thecertificates, use the command:

`(venv)>python gencert.py --final int_cert_data.xlsx`

Check all data in the certificates to verify everything is in order.

A successful run of the script generates a CSV file is written, with the name of the `.xlsx` file suffixed with `_DB` and an extension `.csv`. It containes all data, including certificate number, owner password, which can be stored for subsequent use, if required.

# To Do
Certificate number has the format `YYYY/nnnn`, where `YYYY` is the four digit year during which the certificate is printed and `nnnn` is the four digit number of thecertificate in that year. The year and number of the last certificate printed is read from the `gencert.toml` file. Certificate numbering is continued from the number of the last certificate printed. If the current year, read from the system date is later than the year read from the `gencert.toml` file, the certificate number is reset to `1` and current year is used.

However, the script, at present, does not write the year and certificate number back to `gencert.toml` at the end of printing all the certificates. These values must be changed manually by the user using a text editor.

It is a straight-forward task to take input data from a CSV file instead of a Microsoft Excel `.xlsx` file, as long as the date format is handled correctly.
