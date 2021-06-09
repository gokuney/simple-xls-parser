from modules.xlsProcessor import XLSProcessor
import pdfkit,shutil,os



XLSP = XLSProcessor()

fileName = "./book.xlsx"

# Cleanup
if os.path.exists("output/"):
    shutil.rmtree("output/")
os.mkdir("output/")
        


# Read Excel file
xls = XLSP.readExcel(fileName)

# List sheets
sheets = xls.sheet_names

# Loop sheets
sheetsData = {}
for sheet in sheets:
    data = XLSP.prepareDFs(fileName, sheet)
    # sheetsDF[sheet] = data
    d = XLSP.getProcessedSheet(data, sheet)
    sheetsData[sheet] = d

# Save this data to file
print(sheetsData)

XLSP.saveJSON(sheetsData)


# Create and Zip output

XLSP.createOUTPUT()

XLSP.zipOutput()

XLSP.uploadFile()