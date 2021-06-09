import openpyxl, os,pathlib, shortuuid,shutil,json,requests,webbrowser
from fuzzywuzzy import fuzz
from zipfile import ZipFile
import pandas as pd
from openpyxl_image_loader import SheetImageLoader
from jinja2 import Environment, FileSystemLoader


class XLSProcessor:
    def __init__(self):
        # Clean the output folder
        outputDir = "output/"
        self.outputDir = outputDir
        cwd = str( pathlib.Path().absolute() )
        print("Inited XLS Processor")

    def readExcel(self, path):
        xls = pd.ExcelFile(path)
        return xls
    
    def prepareDFs(self, filePath, sheet):
        df = pd.read_excel(filePath, sheet_name=sheet)
        pxl_doc = openpyxl.load_workbook(filePath)
        sheetData = pxl_doc[sheet]
        return {"df": df, "images": SheetImageLoader(sheetData) }

    def processSheet(self, sheet):
        print(sheet)

    def _allBalnks(self, record):
        resp = True
        for key in record:
            # print("KEY {key} record {record}".format(key=key,record=str(record[key])))
            if str(record[key]) != "nan" or None:
                resp = False
        return resp

    def _getRelevantRowImage(self, images, index):
        fileName = false
        for imageLoc in images._images:
            if str(index) in imageLoc:
                # Save this image
                img = images.get(imageLoc)
                fileName = "%s.jpg" % shortuuid.uuid()
                print("SAVING...")
                img.save(self.outputDir+fileName)
                break
        return fileName

    def _getColumnName(self, toMatch, data, index):
        for column in data.columns:
            if fuzz.ratio(toMatch, column) > 90:
                return data.iloc[index][column] if str(data.iloc[index][column]) != "nan" else "N/A"
        return "N/A"    

    def getProcessedSheet(self, data, sheetName):
        df = data["df"]
        count = 2
        final = []
        images = data["images"]
        for (index,row) in df.iterrows():
            print("Iterating record#%s of sheet %s" % (index+1,sheetName))
            tmp = {
                "Code": self._getColumnName("Code", df, index),
                "Material": self._getColumnName("Material", df, index),
                "Finish":   self._getColumnName("Finish", df, index),
                "Weight":   self._getColumnName("Weight", df, index),
                "Length":   self._getColumnName("Length", df, index),
                "Breadth":  self._getColumnName("Breadth", df, index),
                "Radius":   self._getColumnName("radius", df, index),
                "Rate":     self._getColumnName("Rate", df, index),
                "Image": self._getRelevantRowImage(images, count)
            }
            count = count+1
            if not self._allBalnks(tmp):
                final.append(tmp)

        return final

    def saveJSON(self, data):
        with open(self.outputDir+"output.json", 'w') as file:
            file.write( json.dumps(data) )

    def zipOutput(self):
        shutil.make_archive(os.path.join( '', 'output' ), 'zip', 'output/')

    def createOUTPUT(self):
        data = json.load(  open("output/output.json") )
        sourceTemplate = "templates/light-1"
        fl = FileSystemLoader(sourceTemplate)
        env = Environment(loader=fl)
        template = env.get_template('main.html')
        output = template.render(products=data)
        with open(self.outputDir+"index.html", 'w') as file:
            file.write( output )
        # Copy data
        shutil.copytree("templates/light-1/assets", "output/assets", copy_function = shutil.copy)
        

    def uploadFile(self):
        print("----------------------------")
        print("Uploading file...Please wait")
        print("----------------------------")
        url = 'http://141.164.40.17:3000'
        ui_url = 'http://141.164.40.17'
        files = {'file': open('output.zip', 'rb')}

        r = requests.post(url+"/upload", files=files)
        webbrowser.open(ui_url, new=2)