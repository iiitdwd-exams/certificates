import segno
from PIL import Image
from pathlib import Path
from pikepdf import Pdf, Page, Rectangle


# from certutils import str2path, qr_string


def str2path(fpath: str | Path):
    return Path(fpath) if isinstance(fpath, str) else fpath


def qr_string(data: dict, fields: list[str]):
    qr_str = []
    keys = data.keys()
    for field in fields:
        if field in keys:
            qr_str.append(f"{field.replace("_", " ").capitalize()}: {data[field]}")
    qr_str.append("Issued by: Indian Institute of Information Technology Dharwad")
    return "\n".join(qr_str)


def make_qr(text: str, img_path: str, scale: int = 6, border: int = 2):
    qrcode = segno.make_qr(text)
    qrcode.save(img_path, scale=scale, border=border)


def png2pdf(img_fname: str | Path):
    img_path = str2path(img_fname)
    if img_path.exists():
        img = Image.open(img_path)
        img_rgb = img.convert("RGB")
        pdf_path = img_path.with_suffix(".pdf")
        img_rgb.save(pdf_path)
        img_path.unlink()


def pdf_overlay(
    pdf_cert: str,
    pdf_qr: str | Path,
    pdf_with_qr: str = "",
    x1: int = 65,
    y1: int = 230,
    size: int = 60,
):
    pdf_qr = str2path(pdf_qr)
    if not pdf_with_qr:
        pdf_with_qr = pdf_cert
    allow = pdf_cert == pdf_with_qr
    pdf1 = Pdf.open(pdf_cert, allow_overwriting_input=allow)
    pdf2 = Pdf.open(pdf_qr)
    dest_page = Page(pdf1.pages[0])
    qr_page = Page(pdf2.pages[0])
    x2 = x1 + size
    y2 = y1 + size
    dest_page.add_overlay(qr_page, Rectangle(x1, y1, x2, y2))
    pdf1.save(pdf_with_qr)
    pdf1.close()
    pdf2.close()
    Path(pdf_qr).unlink()


if __name__ == "__main__":
    data = {
        "date": "25-07-2024",
        "student_name": "Ms. Fareah Rahman",
        "supervisor": "Dr. Abdul Wahid",
        "topic": "NPS and AGCPS Algorithms",
    }
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
    make_qr(qr_str, "qr_cert_2024_0001.png")
    png2pdf("qr_cert_2024_0001.png")

    pdf_overlay("cert_2024_0001.pdf", "qr_cert_2024_0001.pdf", "cert_2024_0001_qr.pdf")
