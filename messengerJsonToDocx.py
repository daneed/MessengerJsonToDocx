import os, json, datetime, re
import argparse, pathlib
from docx import Document
from docx.shared import RGBColor, Cm, Pt
from PIL import Image

class Processor (object):
    def __init__(self, file):
        self.file = file
        self.senderNameToColorDict = dict ()
        self.bankAccountRegexObject = re.compile(r'.*([1-9][0-9]{7})([\-|\s]*)([0-9]{8})([\-|\s]*)([0-9]{8}).*', re.IGNORECASE|re.MULTILINE)
        self.phoneNumberRegexObject = re.compile(r'.*(\+36|06)([\s|\-]*)([0-9]{2})([\s|\-]*)([0-9]{3})([\s|\-]*)([0-9]{4}).*', re.IGNORECASE|re.MULTILINE)

    def do(self, prefix, printCallback):
        with open(str(self.file), 'r', encoding="utf-8") as j:
            self.jsonContent = json.loads(j.read())

        if not 'messages' in self.jsonContent or len (self.jsonContent['messages']) == 0:
            print (f'{prefix} File does not contain any messages.')
            return

        document = Document()
        document.add_heading(f"Messenger chat közöttük: {", ".join (self.jsonContent['participants'])}")

        for participant in self.jsonContent['participants']:
            self.senderNameToColor (prefix, participant)

        forbiddenStringRegexpPatterns = self.getForbiddenStringRegexpPatterns(prefix)
        replacements = dict()

        count = 0
        for message in self.jsonContent['messages']:
            printCallback(count, False)
            count += 1
            type = message['type']
            senderName = message['senderName']
            para0 = document.add_paragraph('')
            para0.style.paragraph_format.space_before= Pt(0)
            para0.style.paragraph_format.space_after= Pt(0)
            para0.style.next_paragraph_style = None
            myDate = datetime.datetime.fromtimestamp(int (message['timestamp']/1000.0))
            if message['timestamp'] == 1734784308085:
                pass
            document.add_paragraph (f'[{str(myDate)}]')
            table = document.add_table(rows=1, cols=2)
            table.autofit = True
            nameCell = table.cell(0, 0)
            dataCell = table.cell(0, 1)
            nameCell.width = Cm(3.5)
            dataCell.width = Cm(12.5)
            color = self.senderNameToColor(prefix, senderName)
            senderNameRun = nameCell.paragraphs[0].add_run(f'{senderName}:')
            senderNameRun.font.bold = True
            senderNameRun.font.color.rgb = color
            if 'text' in message:
                messageText = str (message['text'])
                prettyMessageText = messageText.replace(f'\n', f'\n{prefix}')

                for pattern in forbiddenStringRegexpPatterns:
                    match = pattern.search (messageText)
                    if match:
                        if pattern in replacements:
                            replacement = replacements[pattern]
                        else:
                            print (f"\n{prefix}NOTE: FOUND\n{prefix} '{match.group()}' \n{prefix}IN\n{prefix}'{prettyMessageText}'")
                            replacement = input(f'Enter replacement for {match.group()} : ').strip()
                            applyForAllLater = input ('Enter y for apply all later existance: ').upper().strip()
                            if applyForAllLater:
                                replacements[pattern] = replacement

                        if len (replacement) > 0:
                            messageText = re.sub (pattern, replacement, messageText)

                dataCell.paragraphs[0].add_run(f' {messageText}')
            elif type == 'media':
                mediaRun = dataCell.paragraphs[0].add_run()
                mediaRun.font.italic = True
                for media in message['media']:
                    if 'uri' in media:
                        mediaName = media['uri'][2:]
                        picturePath = self.file.absolute().parent/mediaName
                        if (picturePath.suffix in ['.jpeg', '.gif']):
                            im = Image.open(str (picturePath))
                            imWidth, imHeight = im.size
                            width = None
                            height = None
                            if imWidth > imHeight:
                                width = Cm(10)
                            elif imHeight > 500:
                                height = Cm(5)
                            mediaRun.add_picture (str (picturePath), width=width, height=height)
                        elif picturePath.suffix:
                            details = input(f"\nPlease enter details of {mediaName}. You can leave empty: ")
                            if details:
                                mediaRun.add_text (f"<{mediaName}>: {details}")
                            else:
                                mediaRun.add_text (f"<{mediaName}>.")
                        else:
                            mediaRun.add_text (f"<{media['uri']}>")

        docName = (pathlib.Path (os.getcwd()) / self.file.name).with_suffix('.docx')
        document.save (docName)
        printCallback(count, True)
    
    def senderNameToColor (self, prefix, senderName):

        if senderName not in self.senderNameToColorDict:
            color=input(f"{prefix}Enter color for {senderName}. #RRGGBB, or R for red, G for green, B for blue: ")
            color = color.upper()
            if color == 'R':
                self.senderNameToColorDict[senderName] = RGBColor(0xFF, 0x00, 0x00)
            elif color == 'G':
                self.senderNameToColorDict[senderName] = RGBColor(0x00, 0xFF, 0x00)
            elif color == 'B':
                self.senderNameToColorDict[senderName] = RGBColor(0x00, 0x00, 0xFF)
            else:
                if not color.startswith('#'):
                    raise Exception ("Illegal color format!")
                else:
                    try:
                        hexR = int (f"0x{color[1:3]}", 16)
                        hexG = int (f"0x{color[3:5]}", 16)
                        hexB = int (f"0x{color[5:]}", 16)
                        self.senderNameToColorDict[senderName] = RGBColor(hexR, hexG, hexB)
                    except:
                        raise Exception ("Illegal color format!")
            print (f"{prefix}This color will be used for {senderName}: #{str (self.senderNameToColorDict[senderName])}")
        return self.senderNameToColorDict[senderName]
    
    def getForbiddenStringRegexpPatterns (self, prefix):
        canContinue = True
        forbiddenStrings = list()
        while (canContinue):
            forbiddenString = input (f'{prefix}Enter a forbidden string, or press enter to continue: ').strip()
            if len(forbiddenString) > 0:
                forbiddenStrings.append(re.compile (rf'{forbiddenString}', re.IGNORECASE|re.MULTILINE))
            else:
                canContinue=False

        doMask=input (f'{prefix}Max bank accounts? press Y to yes: ').strip().upper()
        if doMask == 'Y':
            forbiddenStrings.append (self.bankAccountRegexObject)

        doMask=input (f'{prefix}Max phone numbers? press Y to yes: ').strip().upper()
        if doMask == 'Y':
            forbiddenStrings.append (self.phoneNumberRegexObject)
        
        return forbiddenStrings

if __name__ == "__main__":
    parser = argparse.ArgumentParser ("Messenger Json to Docx converter")
    parser.add_argument ("-f", "--fileOrFolder", help="The input json file or the container folder", metavar="JSONFILE", required=True)
    args = parser.parse_args()
    path = pathlib.Path(args.fileOrFolder)

    def printFunction (prefix, count, isEnd):
        if isEnd:
            print ("DONE", flush=True)
        elif count == 0:
            print (f"{prefix}Processing messages...", end="", flush=True)
        elif count % 100 == 0:
            print (".", end="", flush=True)

    if path.is_file() and path.suffix == ".json":
        print (f"Processing file: {path.name}...", flush=True)
        processor = Processor(path)
        processor.do ("   ", lambda count, isEnd: printFunction ("   ", count, isEnd))
        print (f"Processing file: {path.name}...DONE", flush=True)

    elif path.is_dir ():
        print (f"Processing directory: {path.name}...")
        for subPath in path.iterdir():
            if subPath.is_file() and subPath.suffix == ".json":
                print (f"   Processing file: {subPath.name}...", flush=True)
                processor = Processor(subPath)
                processor.do (
                    "      ",
                    lambda count, isEnd: printFunction ("   ", count, isEnd))
                print (f"   Processing file: {subPath.name}...DONE", flush=True)
        print (f"Processing directory {path.name}...DONE")
    else:
        raise Exception(f"File {path} is not a json file!")
