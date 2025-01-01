import sys
import os
from os.path import isfile, exists
from datetime import datetime
from pathlib import Path
import uuid
import secrets
import subprocess

import toml
import click
import pandas as pd
from mailmerge import MailMerge

import qr_gen
from pikepdf import Pdf, Encryption, Permissions

if sys.platform == "win32":
    libre_office_path = "c:/Program Files/LibreOffice/program/soffice.exe"
elif sys.platform == "darwin":
    libre_office_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
elif sys.platform == "linux":
    libre_office_path = "/usr/bin/soffice"
else:
    print(f"{sys.platform} not recognised. Program aborted.")
    sys.exit(0)

if not Path(libre_office_path).is_file():
    print(f"Could not find 'LibreOffice' at {libre_office_path}")
    print(
        "Change the path if installed, else install 'LibreOffice'. Then set the path in the script"
    )
    print("Aborting program")
    sys.exit(0)


def get_randomm_filename(prefix: str = "", ext: str = "") -> str:
    return f"{prefix}{uuid.uuid4().hex}{ext}"


def read_df(input_fname: str) -> pd.DataFrame:
    """Read data from a Microsoft Excel or a CSV file and create a DataFrame.
    Returns an empty DataFrame if the data file does not exist
    """
    fpath = Path(input_fname)
    if fpath.exists():
        suffix = fpath.suffix
        if suffix.lower() in [".xls", ".xlsx", ".csv"]:
            if suffix.lower() in [".xls", ".xlsx"]:
                df = pd.read_excel(str(fpath))
            else:
                df = pd.read_csv(str(fpath))
            return df
    else:
        df = pd.DataFrame()
    return df


def mod_name(bare_name: str, gender: str) -> str:
    """Given the bare name (with any prefix Mr., Ms., Ms., Mrs. deleted in advance),
    adds the prefix 'Mr. ' if gender is M, 'Ms. ' if gender is F, else returns unchanged
    """
    match gender[0]:
        case "M":
            return f"Mr. {bare_name}"
        case "F":
            return f"Ms. {bare_name}"
        case _:
            return bare_name


def clean_data(df: pd.DataFrame, prev_cert_num: int):
    """Clean up the DataFrame created previously by reading a Microsoft Excel or a CSV file"""
    df["start_date"] = df["start_date"].dt.strftime("%d-%m-%Y")
    df["end_date"] = df["end_date"].dt.strftime("%d-%m-%Y")
    df["gender"] = (
        df["gender"].str[0].str.upper()
    )  # Retain only the first letter, converted to uppercase
    df["bare_name"] = df["student_name"].str.replace(
        r"^[M]{1}[rs]{1}[s]*\.[ ]*", "", regex=True
    )  # Create a new column with student name stripped of prefixes Mr., Ms. Mrs. and a following space if any
    df["slaut_name"] = df.apply(lambda r: mod_name(r["bare_name"], r["gender"]), axis=1)
    # Create new empty columns
    df["certificate_number"] = ""
    df["certificate_date"] = ""
    df["owner_password"] = ""
    df["download_link"] = ""

    # Converting DataFrame to a list of dict
    start_cert_num = prev_cert_num + 1
    df = df.set_index(pd.Index(range(start_cert_num, start_cert_num + len(df))))
    df_dict = df.to_dict(orient="records")
    return df_dict


def gen_cert_number(year_in_use: int, cert_num: int) -> tuple[int, int]:
    """Generate certificate numbers in the format YYYY/xxxx based on previous values of current_year
    and cert_num. Increments YYYY and rolls over xxxx to 1,  if necesary"""
    current_year = datetime.now().year
    if current_year > year_in_use:
        cert_year = current_year
        cert_num = 1
    else:
        cert_year = year_in_use
        cert_num += 1
    return cert_year, cert_num


def merge_docx(docx_template, output_docx, df_dict):
    """Perform mail merge"""
    with MailMerge(docx_template) as docx:
        docx.merge(**df_dict)
        docx.write(output_docx)


def qr_string(data: dict, fields: list[str]):
    """Generate string to be embedded into QR code based on fields and values printed on the certificate"""
    qr_str = []
    keys = data.keys()
    for field in fields:
        if field in keys:
            qr_str.append(f"{field.replace("_", " ").capitalize()}: {data[field]}")
    qr_str.append("Issued by: Indian Institute of Information Technology Dharwad")
    return "\n".join(qr_str)


def gen_qrpdf(data, qr_fname="qr.png"):
    """Create the string to be embedded into the QR code, save it to PNG format and create a PDF
    with only the QR code
    """
    # Generate the QR code
    fields = [
        "student_name",
        "institute_name",
        "start_date",
        "end_date",
        "supervisor_name",
        "project_title",
        "certificate_date",
        "certificate_number",
    ]
    qr_str = qr_string(data, fields)
    qr_gen.make_qr(qr_str, qr_fname)
    qr_gen.png2pdf(qr_fname)


def encrypt_pdf(
    pdf_infile: str,
    pdf_outfile: str = "",
    owner_password: str = "",
    user_password: str = "",
):
    """Encrypt a PDF file with a owner password with only permission to print.
    Can overwrite the original PDF file (if pdf_outfile is not empty) or
    create a new one (if pdf_outfile is empty)
    """
    if not pdf_outfile:
        pdf_outfile = pdf_infile
        pdf = Pdf.open(pdf_infile, allow_overwriting_input=True)
    else:
        pdf = Pdf.open(pdf_infile)
    permissions = Permissions(
        extract=False,
        modify_assembly=False,
        modify_annotation=False,
        modify_other=False,
        print_highres=True,
        print_lowres=True,
    )
    encryption = Encryption(user=user_password, owner=owner_password, allow=permissions)
    pdf.save(pdf_outfile, encryption=encryption)


def gen_cert(
    data, docx_template: str, docx_cert="cert.docx", pdf_cert="cert.pdf", final=True
):
    """Generates a PDF certificate from a Microsoft Word file with merge fields with values from matching keys
    in a Python dict. Involves the following steps:
      1. Merge fields and save as docx file
      2. Convert merged docx file to PDF using 'soffice'
      3. Create PDF with only the QR code
      4. Overlay the PDF with the QR code on the certificate
      5. Encrypt the overlaid PDF file
    """
    owner_password = ""
    if final:
        tmp_fname = get_randomm_filename()
        tmp_docx_cert = f"{tmp_fname}.docx"
        tmp_pdf_cert = f"{tmp_fname}.pdf"
        # Create docx certificate by merging data in dict into docx template
        merge_docx(docx_template, tmp_docx_cert, data)
        # Export docx certificate to PDF format
        # Requires LibreOffice to be installed, set path to libreoffice suitably
        # Automatically names the pdf file: <filename>.docx -> <filename>.pdf
        subprocess.run(
            [
                libre_office_path,
                "--headless",
                "--convert-to",
                "pdf",
                tmp_docx_cert,
            ]
        )
        # Delete the docx certificate
        Path(tmp_docx_cert).unlink()
        qr_png = f"qr_{tmp_fname}.png"
        gen_qrpdf(data, qr_png)
        # Overlay certificate with QR Code
        qr_gen.pdf_overlay(
            tmp_pdf_cert,
            Path(qr_png).with_suffix(".pdf").name,
            x1=72,
            y1=72,
            size=72,
        )
        # Encrypt PDF
        owner_password = secrets.token_hex(5)
        data["owner_password"] = owner_password
        encrypt_pdf(tmp_pdf_cert, owner_password=owner_password)
        # Rename file
        if exists(pdf_cert):
            os.remove(pdf_cert)
        os.rename(tmp_pdf_cert, pdf_cert)
    return owner_password, pdf_cert


def mangle_name(name: str) -> str:
    """Generate a string with student name mangled for use as part of certificate filename"""
    return (
        name.replace("Mr. ", "")
        .replace("Ms. ", "")
        .lower()
        .replace(". ", "_")
        .replace(".", "_")
        .replace(" ", "_")
    )


@click.command()
@click.option(
    "-d",
    "--date",
    default="",
    help="Date in dd-mm-yyyy format, to be printed on certificate",
)
@click.argument("input_file", type=click.Path(exists=True))
@click.option(
    "--final/--no-final",
    default=False,
    help="Preview, do not actualy create certificates",
)
def main(date, final, input_file):
    db_suffix = "_DB"
    # Input Data
    docx_template = "internship_certificate_template_3.docx"
    # input_file = "internship_certificates_20240826.xlsx"
    if not isfile(docx_template):
        print(f"Template file '{docx_template}' not found. Program aborted")
        sys.exit(1)
    else:
        print(f"Template file: {docx_template}")

    # Config file
    config_file = "gencert.toml"

    if date == "":
        today = datetime.now()  #  datetime(2024, 8, 19) manually set date
    else:
        today = datetime.strptime(date, "%d-%m-%Y")
    cert_date = today.strftime("%d-%m-%Y")
    print(f"Date on certificate: {cert_date}")

    if not isfile(input_file):
        print(f"Certificate data input file '{input_file}' not found. Program aborted")
        sys.exit(2)
    if not isfile(config_file):
        print(
            f"Configuration file '{config_file}' not found. Please enter the following data"
        )
        cert_year = int(input("Year prefix in certificate number: "))
        cert_num = int(input("Number of first certificate to be printed:"))
        cert_num -= 1
    else:
        # Configuration data
        config = toml.load(open(config_file, "r"))
        cert_year = config["certificate"]["year"]
        cert_num = config["certificate"]["cert_num"]
    print(
        f"Continuing from previous data: Year = {cert_year}, Certificate number = {cert_num}"
    )

    # Read and clean data
    df = read_df(input_file)
    print(
        f"Input file: {input_file}. {len(df)} record{'s' if len(df) > 1 else ''} read"
    )
    records = clean_data(df, cert_num)

    print(
        f"{'Number':>9} {'Student Name':30} {'Supervisor':30} {'PDF Password':12} {'Certificate file'}"
    )
    for data in records:
        cert_year, cert_num = gen_cert_number(cert_year, cert_num)
        data["certificate_number"] = f"{cert_year:4d}/{cert_num:04d}"
        data["certificate_date"] = max(cert_date, data["end_date"])
        docx_prefix = data["certificate_number"].replace("/", "_")
        docx_output = Path(
            f"int_cert_{docx_prefix}_{mangle_name(data["student_name"])}.docx"
        )
        pdf_output = Path(docx_output).with_suffix(".pdf").name
        s = f'{data['certificate_number']:9} {data["student_name"]:30} {data["supervisor_name"]:30}'
        print(s, end=" ")
        owner_password, pdf_cert = gen_cert(
            data, docx_template, docx_output.name, pdf_output, final=final
        )
        data["owner_password"] = owner_password

        s = f"{owner_password:12} {pdf_cert}"
        print(s)

    if final:
        pd.DataFrame(records).to_csv(
            Path(input_file).stem + f"{db_suffix}.csv", index=False
        )
        print("Config file updated")
        with open("internship_cert.toml", "w") as f:
            config["certificate"]["year"] = int(cert_year)
            config["certificate"]["cert_num"] = int(cert_num)
            _ = toml.dump(config, f)
        Path("gencert.toml").unlink()
        Path("internship_cert.toml").rename("gencert.toml")
    return


if __name__ == "__main__":
    main()
