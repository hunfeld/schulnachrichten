import os
import csv
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
import PyPDF2
from io import BytesIO
from datetime import datetime

def merge_pages_top_bottom(input_file):
    file = open(input_file, 'rb')
    reader = PyPDF2.PdfFileReader(file)
    writer = PyPDF2.PdfFileWriter()

    page1 = reader.getPage(0)
    page2 = reader.getPage(3)
    new_page = writer.addBlankPage(width=842, height=595)  # DIN A4 size
    new_page.mergeTranslatedPage(page1, 0, 0)
    new_page.mergeTranslatedPage(page2, new_page.mediaBox.getWidth() / 2, 0)
    new_page.rotateClockwise(90)

    page1 = reader.getPage(1)
    page2 = reader.getPage(2)
    new_page = writer.addBlankPage(width=842, height=595)  # DIN A4 size
    new_page.mergeTranslatedPage(page1, 0, 0)
    new_page.mergeTranslatedPage(page2, new_page.mediaBox.getWidth() / 2, 0)
    new_page.rotateClockwise(-90)

    return writer, file

def write_print_instructions(log, num_sheets):
    log.write("\n")
    log.write("Druckanleitung:\n")
    log.write("- Bitte drucken Sie beidseitig.\n")
    log.write("- Die Bindeposition sollte die lange Seite sein.\n")
    log.write("- Bitte prüfen Sie, ob der Druck in Farbe erfolgen soll \n")
    log.write("  oder ob ein Graustufendruck ausreicht.")    
    log.write(f"- Sie benötigen {num_sheets} Blätter.\n\n")

def write_final_instructions(log, file_data):
    log.write("\n")
    log.write("Druckanleitung:\n")
    for filename, num_sheets in file_data.items():
        log.write(f"- Für die Datei {filename} benötigen Sie {num_sheets} Blätter.\n")
    log.write("\nBitte drucken Sie beidseitig und die Bindeposition sollte die lange Seite sein.\n")
    log.write("Bitte prüfen Sie, ob der Druck in Farbe erfolgen soll ")
    log.write("oder ob ein Graustufendruck ausreicht.")   

def write_file_generated_log(log, filename, num_pages):
    log.write(f"{datetime.now().strftime('%H:%M:%S')}: {filename} wurde mit {num_pages} Seiten generiert.\n\n")

with open('Klassen.csv', 'r') as f:
    reader = csv.reader(f, delimiter=';')
    class_data = list(reader)

with open('log.txt', 'w') as log:
    log.write("Herzlichen Dank für die Nutzung des Aufbereiters der Schulnachrichten für den Kopierer.\n\n")
    log.write(f"{datetime.now().strftime('%H:%M:%S')}: Beginn des Durchlaufs.\n\n")

file_data = {}

for input_file in os.listdir():
    if input_file.endswith('.pdf'):
        file = open(input_file, 'rb')
        reader = PyPDF2.PdfFileReader(file)

        if reader.getNumPages() > 20:
            with open('log.txt', 'a') as log:
                log.write(f"{datetime.now().strftime('%H:%M:%S')}: {input_file} hat {reader.getNumPages()} Seiten. Das Maximum von 20 Seiten wird überschritten.\n\n")
            continue
        elif reader.getNumPages() % 4 != 0:
            with open('log.txt', 'a') as log:
                log.write(f"{datetime.now().strftime('%H:%M:%S')}: {input_file} hat {reader.getNumPages()} Seiten. Die Seitenzahl ist kein Vielfaches von 4. Daher wird die Bearbeitung übersprungen.\n\n")
            continue

        file.close()

        output = PdfFileWriter()

        for class_name, class_size in class_data:
            document, file = merge_pages_top_bottom(input_file)
            for _ in range(int(class_size)):
                output.addPage(document.getPage(0))
                output.addPage(document.getPage(1))

            packet = BytesIO()
            querling = canvas.Canvas(packet)
            querling.setFont("Helvetica", 48)  # Setzen Sie die Schriftgröße auf 48
            querling.drawString(100, 500, f"{class_name} ({class_size} Blätter)")  # Fügen Sie den Text hinzu
            querling.showPage()  # Fügen Sie eine leere Rückseite hinzu
            querling.setFont("Helvetica", 48)  # Setzen Sie die Schriftgröße auf 48           
            querling.drawString(100, 500, f"{class_name} ({class_size} Blätter)") 
            querling.save()

            packet.seek(0)
            querling = PdfFileReader(packet)
            output.addPage(querling.getPage(0))
            output.addPage(querling.getPage(1))

        output_file = input_file.replace('.pdf', f'_Kopierdatei_{output.getNumPages()//2}_Blatt.pdf')
        with open(output_file, 'wb') as f:
            output.write(f)

        file_data[output_file] = output.getNumPages()//2

        with open('log.txt', 'a') as log:
            write_file_generated_log(log, output_file, output.getNumPages())

        file.close()

with open('log.txt', 'a') as log:
    write_final_instructions(log, file_data)
