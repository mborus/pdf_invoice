import locale

locale.setlocale(locale.LC_ALL, "German")
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import NameObject, createStringObject

import json
import base64
import tempfile
import shutil
import logging
import os

logging.basicConfig(level=logging.DEBUG)

POINT = 1
INCH = 72.0
CM = INCH / 2.540

PDF_DATA_TAG = "/INVOICE-DATA"


class NoXMLinPdfException(RuntimeError):
    pass


def get_hidden_data_from_pdf(inpdf):

    with open(inpdf, "rb") as p1:
        hidden_data = PdfFileReader(p1).documentInfo.get(PDF_DATA_TAG, "")
        if hidden_data:
            return base64.b64decode(hidden_data).decode("utf-8")
        raise NoXMLinPdfException("No xml in pdf {}".format(inpdf))


def embed_hidden_data_into_pdf(inpdf, indata):
    with open(indata, "r", encoding="cp850") as f1:

        mydata = f1.read()

        # Read xml and encode it
        mydata_enc = base64.b64encode(mydata.encode("utf-8"))
        logging.debug(mydata_enc)
        logging.debug(type(mydata_enc))
        logging.debug(base64.b64decode(mydata_enc).decode("utf-8"))

        with open(inpdf, "rb") as p1:
            tempfile_pdf = tempfile.NamedTemporaryFile(
                mode="w+b", delete=False, suffix=".pdf"
            )

            invoice = PdfFileReader(p1)
            output_pdf = PdfFileWriter()

            infodict = output_pdf._info.getObject()
            for k, v in invoice.documentInfo.items():
                infodict.update({NameObject(k): createStringObject(v)})

            infodict.update(
                {NameObject(PDF_DATA_TAG): createStringObject(mydata_enc.decode("utf-8"))}
            )

            for k, v in invoice.documentInfo.items():
                logging.debug("{} {}".format(k, v))

            for i in range(0, invoice.getNumPages()):
                output_pdf.addPage(invoice.getPage(i))

            # save pdf
            output_pdf.write(tempfile_pdf)
            tempfile_pdf.close()

            persistant_tempfile = tempfile_pdf.name
            logging.info("Using tempfile {}".format(tempfile_pdf.name))

        logging.info("validating temp file")

        validated_data = get_hidden_data_from_pdf(persistant_tempfile)
        logging.debug(validated_data)
        assert mydata == validated_data, "embedded data does not match"
        logging.info("{} {} {}".format(inpdf, indata, persistant_tempfile))

    backup_pdf = "{}.bak".format(inpdf)
    shutil.move(inpdf, backup_pdf)
    os.remove(indata)
    shutil.move(persistant_tempfile, inpdf)
    os.remove(backup_pdf)


def make_invoice_from_json(jsonstring):
    """gets an invoice from a json string

    returns filename

    """
    rechnung = json.loads(jsonstring)
    return make_invoice_pdf(rechnung)


def make_invoice_pdf(rechnung):
    """creates an invoice from json"""
    c = canvas.Canvas(rechnung["dateiname"], pagesize=(21.0 * CM, 29.7 * CM))
    c.setStrokeColorRGB(0, 0, 0)
    c.setFillColorRGB(0, 0, 0)

    c.setFont("Helvetica", 12 * POINT)
    v = 28.0 * CM
    for subtline in rechnung["absender"].split("\n"):
        c.drawString(13.5 * CM, v, subtline)
        v -= 12 * POINT

    c.setFont("Helvetica", 8 * POINT)
    v = 25 * CM
    c.drawString(2.5 * CM, v, rechnung["absender_zeile"])
    c.line(2.5 * CM, v - 0.15 * CM, 11 * CM, v - 0.15 * CM)

    c.setFont("Helvetica", 12 * POINT)

    v = 23 * CM
    for subtline in rechnung["anschrift"].split("\n"):
        c.drawString(2.5 * CM, v, subtline)
        v -= 12 * POINT

    v = 19.0 * CM

    c.drawString(13.5 * CM, v, rechnung["datum"])
    v -= 12 * POINT

    c.setFont("Helvetica-Bold", 12 * POINT)
    v = 18 * CM
    c.drawCentredString(10.5 * CM, v, rechnung["nummer"])
    v -= 12 * POINT

    c.setFont("Helvetica", 12 * POINT)
    v = 17 * CM
    for subtline in rechnung["anschreiben"].split("\n"):
        c.drawString(2.5 * CM, v, subtline)
        v -= 12 * POINT

    v -= 12 * POINT

    c.setFont("Helvetica", 12 * POINT)
    c.drawString(2.5 * CM, v, rechnung["header"][0])
    c.drawString(5 * CM, v, rechnung["header"][1])
    c.drawRightString(18.5 * CM, v, rechnung["header"][2])
    c.line(2.5 * CM, v - 0.15 * CM, 18.5 * CM, v - 0.15 * CM)
    v -= 12 * POINT

    v -= 6 * POINT
    c.setFont("Helvetica", 12 * POINT)
    for subtline in rechnung["zeilen"]:
        c.drawString(2.5 * CM, v, subtline[0])
        c.drawRightString(18.5 * CM, v, str(subtline[2]))
        text = subtline[1]
        for text2 in text.split("\n"):
            c.drawString(5 * CM, v, text2)
            v -= 12 * POINT

        if v < 3.5 * CM:
            c.showPage()
            v = 25 * CM

    if rechnung.get("summe"):
        v += 12 * POINT
        c.line(16.5 * CM, v - 0.15 * CM, 18.5 * CM, v - 0.15 * CM)
        v -= 18 * POINT
        c.drawString(13.5 * CM, v, rechnung["summe"][0])
        c.drawRightString(18.5 * CM, v, rechnung["summe"][1])
        v -= 12 * POINT

    if rechnung.get("mwst"):
        v -= 6 * POINT
        c.drawString(13.5 * CM, v, rechnung["mwst"][0])
        c.drawRightString(18.5 * CM, v, rechnung["mwst"][1])
        c.line(16.5 * CM, v - 0.15 * CM, 18.5 * CM, v - 0.15 * CM)
        v -= 12 * POINT

    if rechnung.get("total"):
        v -= 6 * POINT
        c.drawString(13.5 * CM, v, rechnung["total"][0])
        c.drawRightString(18.5 * CM, v, rechnung["total"][1])
        c.line(16.5 * CM, v - 0.15 * CM, 18.5 * CM, v - 0.15 * CM)
        c.line(16.5 * CM, v - 0.25 * CM, 18.5 * CM, v - 0.25 * CM)

    c.setFont("Helvetica", 10 * POINT)
    v = 2.5 * CM
    for subtline in rechnung["fusszeile"].split("\n"):
        c.drawString(2.5 * CM, v, subtline)
        v -= 12 * POINT

    c.showPage()
    c.save()

    return rechnung["dateiname"]


def add_background_to_pdf(
    filename_in,
    filename_out=None,
    filename_letterhead=None,
    filename_background=None,
    new_title="",
    new_author="",
):
    """Merges a one page letterhead to an invoice and sets author and title of the doc info

    for multi page pdfs, its possible to define an extra page

    """

    if not filename_letterhead:
        return

    use_tmpfile = False

    if not filename_out:
        filename_out = tempfile.NamedTemporaryFile(
            mode="w+b", delete=False, suffix=".pdf"
        ).name
        use_tmpfile = True

    if not filename_background:
        filename_background = filename_letterhead

    with open(filename_in, "rb") as pdf_in, open(
        filename_background, "rb"
    ) as pdf_lb, open(filename_letterhead, "rb") as pdf_lh:

        input_pdf = PdfFileReader(pdf_in)
        output_pdf = PdfFileWriter()

        # metadata
        # noinspection PyProtectedMember
        infodict = output_pdf._info.getObject()
        for k, v in input_pdf.documentInfo.items():
            infodict.update({NameObject(k): createStringObject(v)})
        infodict.update({NameObject("/Title"): createStringObject(new_title)})
        infodict.update({NameObject("/Author"): createStringObject(new_author)})

        # add first page
        # get the first invoice page, merge with letterhead
        letterhead = PdfFileReader(pdf_lh).getPage(0)
        letterhead.mergePage(input_pdf.getPage(0))
        output_pdf.addPage(letterhead)
        # add other pages
        for i in range(1, input_pdf.getNumPages()):
            background = PdfFileReader(pdf_lb).getPage(0)
            background.mergePage(input_pdf.getPage(i))
            output_pdf.addPage(background)
        # save pdf
        with open(filename_out, "wb") as pdf_out:
            output_pdf.write(pdf_out)

    if use_tmpfile:
        backup_pdf = "{}.bak".format(filename_in)
        shutil.move(filename_in, backup_pdf)
        shutil.move(filename_out, filename_in)
        os.remove(backup_pdf)


def beispielrechnung(
    dateiname,
    firma="PYBC",
    rechnungsnummer="2019-03",
    datum="13.04.2019",
    zeitraum="März 2019 bis April 2019",
):
    rechnung = {}

    rechnung["dateiname"] = dateiname
    rechnung[
        "absender"
    ] = """
Martin Borus
Fakestrasse 1
24937 Flensburg
Telefon 0461-000000
Email: pythonbc@borus.de

USt-ID DE0000000000
"""
    rechnung["absender_zeile"] = """Martin Borus, Fakestrasse 1, 24937 Flensburg"""

    if firma == "PYBC":
        rechnung["anschrift"] = (
            "GFU Cyrus AG\nAm Grauen Stein 27\n" "51105 Köln-Deutz\n"
        )
        rechnung["anschreiben"] = (
            "Hiermit berechne ich Ihnen wie vereinbart folgende Leistungen für den\n"
            "Zeitraum {} für die Unterstützung bei der\n"
            "Lösung von Problemen, die man ohne Computer gar nicht gehabt hätte \n"
            "zum Stundensatz von €8.84.\n"
        ).format(zeitraum)

    rechnung["nummer"] = "Rechnung {}".format(rechnungsnummer)
    rechnung["datum"] = "Flensburg, den {}".format(datum)

    rechnung["fusszeile"] = (
        "Bitte überweisen Sie den Rechungsbetrag ohne Abzug innerhalb der nächsten 14 Tage auf mein Konto\n"
        "bei der Nospa Flensburg. IBAN: DE79 0000 0000 0000 0000 00, BIC: XXXXXXXXXXX"
    )
    rechnung["header"] = ("Datum", "Beschreibung", "Betrag")
    rechnung["zeilen"] = [
        ("01.03.2019", "7 Arbeitsstunden je €8.84/Stunde", "€61.88"),
        ("02.03.2019", "3 Arbeitsstunden je €8.84/Stunde", "€26.52"),
        ("01.04.2019", "1 Arbeitsstunde je €8.84/Stunde", "€8.84"),
    ]
    rechnung["summe"] = ("Summe:", "€97,24")
    rechnung["mwst"] = ("19% MwSt:", "€13,08")
    rechnung["total"] = ("Gesamt:", "€110.33")
    return rechnung


def create_example():
    """Create an example invoice and save it"""

    inv = "rechnung_nr_2019-04_PYBC_2019-04.json"
    pdf = "rechnung_nr_2019-04_PYBC_2019-04.pdf"

    rechnung = beispielrechnung(
        pdf,
        firma="PYBC",
        zeitraum="März 2019 bis April 2018",
        rechnungsnummer="2019-04",
    )

    with open(inv, "w", encoding="utf-8") as f1:
        f1.write(json.dumps(rechnung))

    return inv


if __name__ == "__main__":

    invoice_json_example = create_example()

    with open(invoice_json_example, encoding="utf-8") as f1:
        invoice_pdf = make_invoice_from_json(f1.read())

    add_background_to_pdf(
        filename_in=invoice_pdf,
        filename_letterhead="sample.pdf",
        filename_background=None,
        new_author="new",
        new_title="new",
    )

    embed_hidden_data_into_pdf(inpdf=invoice_pdf, indata=invoice_json_example)

    logging.info(f"Rechnung erstellt: {invoice_pdf}")

    logging.info('Entschlüsselte Daten')
    logging.info((get_hidden_data_from_pdf(invoice_pdf)))