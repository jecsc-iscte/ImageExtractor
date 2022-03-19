from docx import Document
import requests
from bs4 import BeautifulSoup
import os
import comtypes.client
from PIL import Image
import glob
import tkinter as tk

from docx.shared import Cm


def main():
    gui = tk.Tk()

    label1 = tk.Label(gui, text="Localizacao no sistema:")
    label2 = tk.Label(gui, text="Capitulo:")
    label3 = tk.Label(gui, text="Exemplo : C:\\Users\\Joao\\Documents\\Berserk")
    label4 = tk.Label(gui, text="")
    label1.grid(row=0, column=0, sticky=tk.W, pady=2)
    label2.grid(row=3, column=0, sticky=tk.W, pady=2)
    label3.grid(row=1, column=0, sticky=tk.W, pady=2)
    label4.grid(row=2, column=0, sticky=tk.W, pady=2)

    entry1 = tk.Entry(gui)
    entry1.config(width=50)
    entry2 = tk.Entry(gui)
    entry1.grid(row=0, column=1, pady=2)
    entry2.grid(row=3, column=1, pady=2)

    button = tk.Button(gui, text="Generate", command=lambda: imageExtractor(entry1.get(), entry2.get(), button, label4))
    button.grid(row=4, column=0, pady=2)

    gui.mainloop()

def imageExtractor(folderPATH, chapter, button, label):
    changeButtonState(button, "disable")
    if not os.path.isdir(folderPATH):
        editLabelText(label, "O diretorio nao existe")
    elif not websiteExists(chapter):
        editLabelText(label, "Esse capitulo nao existe")
    else:
        os.chdir(folderPATH)
        editLabelText(label, "Clean das imagens desnecessarias")
        removeDirFiles('jpg', folderPATH, chapter)
        editLabelText(label, "Download das imagens necessarias")
        downloadImages('http://berserkmanga.net/manga/berserk-chapter-'+chapter+'/')
        editLabelText(label, "A criar word")
        addImagesToWord(folderPATH, chapter)
        editLabelText(label, "A converter word para pdf")
        wordToPDF(folderPATH, chapter)
        editLabelText(label, "Feito!")
        removeDirFiles('jpg', folderPATH, chapter)
        editLabelText(label, "")
    changeButtonState(button, "normal")

def changeButtonState(button, state):
    button["state"] = state

def websiteExists(chapter):
    response = requests.get('http://berserkmanga.net/manga/berserk-chapter-'+chapter+'/')
    if response.status_code == 200:
        return True
    else:
        return False

def editLabelText(label, newText):
    label.config(text=newText, highlightcolor="red")

def downloadImages(url):
    r = requests.get(url)
    soup = BeautifulSoup(r.text, 'html.parser')

    images = soup.find_all('img')
    i = 1100
    for image in images:
        link = image['src']
        with open("tempBerserk" + str(i) + '.jpg', 'wb') as f:
            im = requests.get(link)
            f.write(im.content)
        i = i + 1

def addImagesToWord(folderPATH, chapter):
    document = Document()
    updateMargins(document, folderPATH, chapter)
    document = Document(os.path.join(folderPATH, 'capitulo' + chapter + '.docx'))
    for path in glob.glob1(folderPATH, 'tempBerserk*.jpg'):
        img = Image.open(os.path.join(folderPATH, path))
        if img.size[0] > 1500:
            img = img.transpose(Image.ROTATE_270)
            img = img.resize((int(img.size[0] * 0.35), int(img.size[1] * 0.35)), Image.ANTIALIAS)
        else:
            img = img.resize((int(img.size[0] * 0.5), int(img.size[1] * 0.5)), Image.ANTIALIAS)
        img.save(os.path.join(folderPATH, path))
        p = document.add_paragraph()
        r = p.add_run()
        r.add_picture(os.path.join(folderPATH, path))
    document.save(os.path.join(folderPATH, 'capitulo' + chapter + '.docx'))

def updateMargins(document, folderPATH, chapter):
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(0)
        section.left_margin = Cm(0.8)

    document.save(os.path.join(folderPATH, 'capitulo' + chapter + '.docx'))

def wordToPDF(folderPATH, chapter):
    in_file = os.path.abspath(os.path.join(folderPATH, 'capitulo' + chapter + '.docx'))
    out_file = os.path.abspath(os.path.join(folderPATH, 'capitulo' + chapter + '.pdf'))

    wdFormatPDF = 17

    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

def removeDirFiles(fileType, folderPATH, chapter):
    for path in glob.glob1(folderPATH, 'tempBerserk*.'+fileType):
        os.remove(os.path.join(folderPATH, path))
    if os.path.isfile(os.path.join(folderPATH, 'capitulo' + chapter + '.docx')):
        os.remove(os.path.join(folderPATH, 'capitulo' + chapter + '.docx'))

if __name__ == '__main__':
    main()
