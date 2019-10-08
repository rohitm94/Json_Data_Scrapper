import json
from glob import glob
import xlsxwriter

workbook = xlsxwriter.Workbook('<excel file name>.xlsx')
worksheet = workbook.add_worksheet()

bold = workbook.add_format({'bold': 1})     #Styling the excel inputs

worksheet.write('A1', 'Dataset_id', bold)
worksheet.write('B1', 'Publication Date', bold)
worksheet.write('C1', 'id', bold)
worksheet.write('D1', 'Version No.', bold)
worksheet.write('E1', 'Last Update Time', bold)
worksheet.write('F1', 'Release time', bold)
worksheet.write('G1', 'Create time', bold)
worksheet.write('H1', 'License', bold)
worksheet.write('I1', '# Fields', bold)
worksheet.write('J1', 'title', bold)
worksheet.write('K1', 'Author Name', bold)
worksheet.write('L1', 'DataSet Contact Name', bold)
worksheet.write('M1', 'DataSet Contact Affiliation', bold)
worksheet.write('N1', 'DataSet Contact Email', bold)
worksheet.write('O1', 'Subject', bold)
worksheet.write('P1', 'Date of Deposit', bold)
worksheet.write('Q1', '#Files', bold)
worksheet.write('R1', 'Filesize(bytes)', bold)

#Traversing through all the json files in the local folder
file_list = glob('C:/vA/MetaData_Json/scrap/*.json')

row =1
for file in file_list:
    inputFile = open(file,"r",encoding="utf8")
    data = json.load(inputFile)
    worksheet.write_number(row,0,data['data']['id'])
    worksheet.write_string(row,1,data['data']['publicationDate'])
    worksheet.write_number(row,2,data['data']['latestVersion']['id'])
    if "versionNumber" in data['data']['latestVersion']:
        worksheet.write_number(row,3,data['data']['latestVersion']['versionNumber'])
    worksheet.write_string(row,4,data['data']['latestVersion']['lastUpdateTime'])
    worksheet.write_string(row,5,data['data']['latestVersion']['releaseTime'])
    worksheet.write_string(row,6,data['data']['latestVersion']['createTime'])
    if "license" in data['data']['latestVersion']:
        worksheet.write_string(row,7,data['data']['latestVersion']['license'])
    fields = data['data']['latestVersion']['metadataBlocks']['citation']['fields']
    worksheet.write_number(row,8,len(fields))
    for field in fields :
        if(field['typeName']=="title"):
            worksheet.write_string(row,9,field['value'])
        elif(field['typeName']=="author"):
            if(len(field['value'])):
                list_auth=""
                for auth in field['value']:
                    if "authorName" in auth:
                        list_auth = auth['authorName']['value']+"|"+list_auth
                        worksheet.write_string(row,10,list_auth) 
        elif(field['typeName']=="datasetContact"):
            if(len(field['value'])):
                list_contact=""
                list_contactAffi=""
                list_contactEmail=""
                for contact in field['value']:
                    if "datasetContactName" in contact:
                        list_contact = contact['datasetContactName']['value']+"|"+list_contact 
                        worksheet.write_string(row,11,list_contact)
                    if "datasetContactAffiliation" in contact:
                        list_contactAffi = contact['datasetContactAffiliation']['value']+"|"+list_contactAffi
                        worksheet.write_string(row,12,list_contactAffi)
                    if "datasetContactEmail" in contact:
                        list_contactEmail = contact['datasetContactEmail']['value']+"|"+list_contactEmail 
                        worksheet.write_string(row,13,list_contactEmail)  
        elif(field['typeName']=="subject"):
            if(len(field['value'])):
                subject = ""
                for sub in field['value']:
                    subject = sub+":"+subject;
                    worksheet.write_string(row,14,subject) 
        elif(field['typeName']=="dateOfDeposit"):
            worksheet.write_string(row,15,field['value'])
    files = data['data']['latestVersion']['files']
    TotalFiles = 0
    Totalsize = 0
    for datafile in files :
        TotalFiles +=1
        Totalsize = Totalsize + datafile['dataFile']['filesize']
    
    worksheet.write_number(row,16,TotalFiles)
    worksheet.write_number(row,17,Totalsize)
    row+=1

workbook.close()
