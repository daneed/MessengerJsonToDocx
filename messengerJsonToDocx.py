import os, json, datetime
import argparse, pathlib
from docx import Document
from docx.shared import RGBColor, Cm, Pt
from PIL import Image

class Processor (object):
    def __init__(self, file):
        self.file = file

    def do(self):
        with open(str(file), 'r', encoding="utf-8") as j:
            self.jsonContent = json.loads(j.read())
        document = Document()
        document.add_heading(f"Messenger chat közöttük: {", ".join (self.jsonContent['participants'])}")
        for message in self.jsonContent['messages']:
            type = message['type']
            senderName = message['senderName']
            para0 = document.add_paragraph('')
            para0.style.paragraph_format.space_before= Pt(0)
            para0.style.paragraph_format.space_after= Pt(0)
            para0.style.next_paragraph_style = None
            myDate = datetime.datetime.fromtimestamp(int (message['timestamp']/1000.0))
            document.add_paragraph (f'[{str(myDate)}]')
            table = document.add_table(rows=1, cols=2)
            table.autofit = True
            nameCell = table.cell(0, 0)
            dataCell = table.cell(0, 1)
            nameCell.width = Cm(3)
            dataCell.width = Cm(13)
            color = self.senderNameToColor(senderName)
            senderNameRun = nameCell.paragraphs[0].add_run(f'{senderName}:')
            senderNameRun.font.bold = True
            senderNameRun.font.color.rgb = color
            if type == 'text':
                messageText = message['text']
                messageTextRun = dataCell.paragraphs[0].add_run(f' {messageText}')
            elif type == 'media':
                imageRun = dataCell.paragraphs[0].add_run()
                for media in message['media']:
                    if 'uri' in media:
                        picturePath = file.absolute().parent/media['uri'][2:]
                        if (picturePath.suffix == '.jpeg'):
                            im = Image.open(str (picturePath))
                            imWidth, imHeight = im.size
                            width = None
                            height = None
                            if imWidth > imHeight:
                                width = Cm(10)
                            elif imHeight > 500:
                                height = Cm(5)
                            imageRun.add_picture (str (picturePath), width=width, height=height)
        docName = (pathlib.Path (os.getcwd()) / file.name).with_suffix('.docx')
        document.save (docName)
    
    def senderNameToColor (self, senderName):
        if senderName == 'Márk Héger':
            return RGBColor(0xFF, 0x00, 0x00)
        elif senderName == 'Daniel Serly':
            return RGBColor(0x00, 0x00, 0xFF)
        else:
            return RGBColor(0x00, 0x00, 0x00)

if __name__ == "__main__":
    parser = argparse.ArgumentParser ("Messenger Json to Docx converter")
    parser.add_argument ("-f", "--file", help="The input json file", metavar="JSONFILE", required=True)
    args = parser.parse_args()
    file = pathlib.Path(args.file)
    if not file.is_file() or file.suffix != ".json":
        raise Exception(f"File {file} is not a json file!")
    processor = Processor(file)
    processor.do ()
