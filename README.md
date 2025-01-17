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

1. Choose the root directory where you wish to clone the GitHub repo
2. Clone the Github repository `git clone https://github.com/satish-annigeri/certificates.git`.
3. Change into the directory

## Using `uv`
To install `uv`, see instructions here: [Installing `uv`](https://docs.astral.sh/uv/getting-started/installation/). Alternately, you can install `uv` using `pipx`. See the next section for instructions to install `pipx`.

4. Create the virtual environment and install the required packages with the single command `uv sync`

## Using `pip-compile` and `pip-sync`
Install, it is best to first install `pipx` using the installer for your operating system. See here for instructions: [Installing `pipx`](https://pipx.pypa.io/latest/installation/). After installing `pipx`, install `pip-tools` using `pipx` with the command `pipx install pip-tools`. Check to verify that `pip-compile` and `pip-sync` are available.

4. Create a virtual environment inside the newly cloned directory with the command `python -m venv .venv`. The name of the virtual environment created by this command is `.venv`.
5. Activate the virtual environment. On Microsoft Windows: `.venv\Scripts\activate`. On GNU/Linux: `soource .venv/bin/activate`.
6. Create the `requirements.txt` file from the `requirements.in` file already present in the cloed directory with the command: `pip-compile requirements.in -o requirements.txt`
7. Install the packages with the command: `pip-sync  requirements.txt`

You could do the same using `uv` instead of using `pip-tools` as the commands `pip-compile` and `pip-sync` are built-in into `uv`. The equivalent `uv` commands are `uv pip compile` instead of `pip-compile` and `uv pip sync` instead of `pip-sync`.

# Configuration File
The application depends on a configuration file in TOML format named `gencert.toml` in the same directory as the application. 

The format of the configuration file is as follows:
```toml
[certificate]
year = 2025
cert_num = 0
```
The field `year` represent the four digit year and `cert_num` represents the number of the most recent certificate generated. Before the first run of the application, be sure to set `year` to the current year and `cert_num = 0` when you have not generated any certificates previously. This file is overwritten at the end of each run of the application. The `year` and `cert_num` values can be manually changed by the user, if needed.

If the current year is greater than the year in the configuration file, the year value is set to the current year and certificate number is reset to 1.

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
Check if LibreOffice is installed, and if not, install it from here: `https://www.libreoffice.org/`. Determine the path where `soffice` is installed. On GNU/Linux, use the `which soffice` command and on Windows, use `where soffice` command. If it is already on the `PATH` environment variable, it will be located. Else, use any method to determine the path to `soffice`, such as, looking up the properties of LibreOffice in your Windows Start menu. Open the script `gencert.py` in your IDE and search for `libre_office_path` and change it appropriately.

Place the input data file in the same directory as the script and execute the command:

`(venv)>python gencert.py --no-final int_cert_data.xlsx`

where `int_cert_data.xlsx` is the Microsoft Excel file containing the data as described above. With the `--no-final` switch, only a dry run is executed, certificates will not be generated. The screen output gives you an idea if the data was arranged correctly in the input data file.

To generate thecertificates, use the command:

`(venv)>python gencert.py --final int_cert_data.xlsx`

Check all data in the certificates to verify everything is in order.

A successful run of the script generates a CSV file is written, with the name of the `.xlsx` file suffixed with `_DB` and an extension `.csv`. It containes all data, including certificate number, owner password, which can be stored for subsequent use, if required.

# To Do
1. Store the data of all previously generated certificates in a persistent store. At present, data of each run are saved to a CSV file (with the name of the input file suffixed by `_DB`) and an aggregate list must be concatenated manually.
2. Store data in a database, perhaps in an SQLite 3 database, instead of in a `.xlsx` file.
3. Convert the application from a CLI to a Streamlit app.
4. Use gender to automatically generate "Mr." or "Ms. when printing student name on certificate (at present, this must be included with the name of student during input).
