import pandas as pd
from openpyxl import load_workbook

# replace file names here
input_filename = 'Draft CCS and H data framework - public support.xlsx'
output_filename = 'merged.xlsx'

# tab names, rows to be skipped based on tab formatting
h2 = pd.read_excel(input_filename, sheet_name='h2ccs')
ccus = pd.read_excel(input_filename, sheet_name='ccus')
final = pd.read_excel(input_filename, sheet_name='final', skiprows=[0, 2, 3, 4, 5])

# dictionary mapping column names in different sheets
ccus_header_dict = {'ProjectStage': 'Project status',
                    'FY': 'Financing year',
                    'RecipientEntity': 'Companies involved',
                    'country': 'Country',
                    'projectType': 'Project type',
                    'projectScale': 'Project scale',
                    'subsector': 'Subsector',
                    'typeOfCapture': 'Type of capture',
                    'captureTech': 'Capture tech',
                    'retroFit': 'Retrofit',
                    'co2Source': 'CO2 source',
                    'co2Destination': 'Application (CO2 destination)',
                    'distanceToStorage': 'Distance to storage (km)',
                    'co2capture (Mt p.a.)':'Capture capacity (MtCO2/yr)'}

def update_final_sheet(final, ccus, h2):
  '''Update final sheet entries with corresponding entries in bloomberg sheet, if applicable'''
  # join spreadsheets along the name of the bloomberg entry
  df = final.join(ccus.set_index('Project Name'), on='bloombergName')
  # replace all empty entries in final with corresponding entries in bloomberg sheet
  for key, value in ccus_header_dict.items():
    final[key].fillna(df[value], inplace=True)

def save_sheet(df, input_filename, output_filename, sheet_name='merged'):
  '''Save dataframe as sheet of exisiting excel workbook'''
  book = load_workbook(input_filename)
  writer = pd.ExcelWriter(output_filename, engine='openpyxl') 
  writer.book = book

  ## ExcelWriter for some reason uses writer.sheets to access the sheet.
  ## If you leave it empty it will not know that sheet Main is already there
  ## and will create a new sheet.
  writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

  # save merged data to additional sheet
  df.to_excel(writer, sheet_name=sheet_name)
  writer.save()

if __name__ == "__main__":
  update_final_sheet(final, ccus, h2)
  save_sheet(final, input_filename=input_filename, output_filename=output_filename, sheet_name='merged')