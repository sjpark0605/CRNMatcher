from tkinter import Tk, Label, Frame, Button
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from Logic import matchCRN

class GUI:
    def __init__(self):
        self.root = Tk()
        self.root.title("테스트 프로그램")
        self.root.resizable(False, False)
        self.root.geometry("600x110")
        self.pathMap = {}

    def run(self):
        self.addField("사업자 정보 엑셀", 0, 0)
        self.addField("통신판매업 엑셀", 0, 1)
        self.addField("저장할 폴더", 1, 2)
        self.addProcessButton()
        self.root.mainloop()

    def addProcessButton(self):
        action = lambda: self.mergeExcelFiles()
        mergeButton = self.generateButton("엑셀 병합", action)
        mergeButton.grid(row = 3, column = 0)


    def mergeExcelFiles(self):
        try:
            excel1 = self.readPath("사업자 정보 엑셀")
            excel2 = self.readPath("통신판매업 엑셀")
            outputDir = self.readPath("저장할 폴더")

            matchCRN(excel1, excel2, outputDir)

            showinfo(
                title = "병합 성공",
                message = "엑셀이 " + outputDir + "/통판사업자정보.xlsx로 저장 되었습니다."
            )

        except Exception as e:
            showinfo(
                title = "오류",
                message = str(e)
            )

    def readPath(self, label):
        path = self.pathMap[label].cget("text")

        if path != None and path != "":
            return self.pathMap[label].cget("text")
        else:
            if label == "저장할 폴더":
                raise Exception(label + "를 지정하세요")
            else: 
                raise Exception(label + "을 선택하세요")

    def addField(self, labelContent, mode, row):
        fieldLabel = self.generateFieldLabel(labelContent)
        pathFrame = self.generateFrame()

        pathLabel = self.generatePathLabel(pathFrame)
        self.pathMap[labelContent] = pathLabel

        if mode == 0:
            action = lambda: self.getFilePath(labelContent + " 불러오기", pathLabel)
            button = self.generateButton("불러오기", action)
        else:
            action = lambda: self.getDirPath(pathLabel)
            button = self.generateButton("폴더 선택", action)

        self.renderItems(fieldLabel, pathFrame, pathLabel, button, row)

    def renderItems(self, fieldLabel, pathFrame, pathLabel, button, row):
        fieldLabel.grid(row = row, column = 0)
        pathFrame.grid(row = row, column = 1)
        pathLabel.pack()
        button.grid(row = row, column = 2)

    def generateFieldLabel(self, labelContent):
        fieldLabel = Label(
            self.root,
            text = labelContent,
            padx = 5,
            pady = 3
        )

        return fieldLabel

    def generateFrame(self):
        frame = Frame(
            self.root,
            bg="black",
            padx = 1,
            pady = 1
        )

        return frame

    def generatePathLabel(self, frame):
        pathLabel = Label(
            frame, 
            bg = "white",
            width = 50
        )

        return pathLabel

    def getDirPath(self, pathLabel):
        targetDir = fd.askdirectory(
            title = "저장할 위치 선택"
        )

        pathLabel.config(text = targetDir)
        

    def getFilePath(self, filePrompt, pathLabel):
        filetypes = (
            ('엑셀 파일', '*.xlsx'),
            ('모든 파일', '*.*')
        )

        filename = fd.askopenfilename(
            title = filePrompt,
            initialdir = '/',
            filetypes = filetypes
        )

        pathLabel.config(text = filename)

    def generateButton(self, text, action):
        button = Button(
            self.root, 
            text = text,
            command = action,
            padx = 10
        )

        return button
        
