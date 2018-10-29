
# coding: utf-8

import os
import time
import datetime
import warnings
import pandas as pd
import numpy as np
import hashlib
import json
import pickle
import traceback

is_getExcel=False
is_getWord=False
is_getPdf=True
is_pdf2png = False
is_getMsg=False

def _GetSubmissionNumber(filename):
    SubmissionNumber = filename.split('\\')[-2]
    return SubmissionNumber
class CreateJSON():
    def __init__(self):
        pass
    def _GetJsonFileName(self,filename):
        file_format = filename.split('.')[-1].lower()
        File = "".join(filename.split("\\")[-1].split('.')[:-1])
        filepath ='\\'.join(filename.split('\\')[:-1])
        FileDirectory = '\\'.join(filename.split('\\')[-6:-1])
        JsonFolder = SharedFolderInitial+":\\Starr Sprint-F18\\Data\\JSON Output\\"
        new_directory=os.path.join(JsonFolder,FileDirectory)
#         print(new_directory)
        if not os.path.exists(new_directory):
            os.makedirs(new_directory)
        JsonFileName = File +"_"+ file_format +".json"
        return os.path.join(new_directory,JsonFileName)    
    def ToJson(self,filename,extraction):     
        JsonFileName = self._GetJsonFileName(filename)
        with open(JsonFileName, 'w') as outfile:
            json.dump(extraction,outfile)
            print(JsonFileName + ' Processed.')
            print()
class ExcelParser():
    def __init__(self):
        self.CreateJSON=CreateJSON()
    def _num_to_col_letters(self, num):
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))

    def _extract_Excel(self, filename): # Excel Filename, end with xls/xlsx/csv
        from win32com.client import Dispatch
        try:
            wb=None
            excel = Dispatch('Excel.Application')
            excel.visible = 0
            wb = excel.Workbooks.Open(os.path.abspath(filename),UpdateLinks=0)
            wb.CheckCompatibility=False 
            excel_texts = {}       
            for worksheet in excel.Worksheets:
                if worksheet.UsedRange() is None:
                    continue
                pagename = worksheet.Name
                excel_text = [list(L) for L in worksheet.UsedRange()]
                excel_texts[pagename] = {}
                for i in range(len(excel_text)):
                    for j in range(len(excel_text[i])):
                        if excel_text[i][j] == None:
                            continue
                        else: 
                            excel_texts[pagename][str(excel_text[i][j])]=str(i+1)+","+self._num_to_col_letters(j+1)
            wb.Close(True)
            del wb
            del excel
            return excel_texts
        except:
            wb.Close(True) if wb is not None else False
            del wb
            del excel
            traceback.print_exc()

    def _content_extraction_excel(self, excel_texts):
        content_extracted = []
        for page in excel_texts:
            for cell in excel_texts[page]:
                content_extracted.append(str(cell))
        content_extracted = ' '.join(content_extracted)
        return content_extracted 
    def _ParseExcelFile(self, filename):
        if self._extract_Excel(filename) is not None:
            start_time = time.time()
            excel_texts = self._extract_Excel(filename)
            extraction = {}  
            extraction['Excel_Response'] = excel_texts
            content_extracted = self._content_extraction_excel(excel_texts)
            extraction['Content_Extracted'] = content_extracted
            extraction['SubmissionNumber']=_GetSubmissionNumber(filename)
            extraction['FileType']=filename.split("\\")[-1].split('.')[-1].lower()
            extraction['FileID']=hashlib.md5(str(extraction).encode('utf-8')).hexdigest()
            extraction['FileLocation']=filename
            elapse_time = time.time() - start_time
            extraction['TotalElapsedTimeMs'] = str(elapse_time/60.0)  
            date_time = datetime.datetime.utcnow()
            extraction['ProcessedDateTime'] = str(date_time)
            return extraction    
    def dump2Json(self,filename):  
        JsonFileName = self.CreateJSON._GetJsonFileName(filename)
        if not os.path.exists(JsonFileName):
            extraction = self._ParseExcelFile(filename)
            if extraction is None:
                raise ValueError("Extraction is None")
            self.CreateJSON.ToJson(filename,extraction)


class OutlookParser():
    def __init__(self):
        self.CreateJSON=CreateJSON()
    def _GetSenderEmail(self,MailItem):
        if MailItem.SenderEmailType=='EX':
            return MailItem.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            return MailItem.SenderEmailAddress
    def _GetRecipientsEmail(self,MailItem):
        Recipients=[]
        for Recipient in MailItem.Recipients:
            RecipientEmail = Recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
            Recipients.append(RecipientEmail)
        return Recipients    
#     def _GetAttachmentJson(self,Attachment):
#         file_format = Attachment.split('.')[-1].lower()
#         if file_format in ['xls', 'xlsx', 'csv']:
#             JsonFileName = self._GetJsonFileName(Attachment,file_format)
#             if not os.path.exists(JsonFileName):
#                 extraction = self.ParseExcelFile(Attachment)
#                 if extraction is None:
#                     raise ValueError("Extraction is None")
#                 self._ToJson(JsonFileName,extraction)
#         elif file_format in ['doc', 'docx','pdf']:
#             JsonFileName = self._GetJsonFileName(Attachment,file_format)
#             if not os.path.exists(JsonFileName):
#                 extraction = self.Parse_PDF_Doc_File(Attachment)
#                 if extraction is None:
#                     raise ValueError("Extraction is None")
#                 self._ToJson(JsonFileName,extraction)
#        elif file_format in ['msg']:
#            AttIsMsg.append(Attachment)
#             JsonFileName = self._GetJsonFileName(Attachment,file_format)
#             if not os.path.exists(JsonFileName):
#                 extraction = self.ParseOutlookFile(Attachment)
#                 self._ToJson(JsonFileName,extraction)
    def _ParseOutlookFile(self,filename):
        import win32com
        try:
            msg = None
            start_time = time.time()
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            msg = outlook.OpenSharedItem(os.path.abspath(filename))
            MsgFileName = ''.join(filename.split("\\")[-1].split('.')[:-1])
            MsgFileRoot = '\\'.join(filename.split('\\')[-6:-1])
            MsgAttachmentFolder = SharedFolderInitial+':\\Starr Sprint-F18\\Data\\360-documents-MsgAttachment\\'
            MsgAttachment_directory=os.path.join(MsgAttachmentFolder,MsgFileRoot)
            extraction = {}
            MsgMetaData={}
            MsgMetaData['SentTime']=str(msg.SentOn)
            MsgMetaData['Sender']={"Name":msg.SenderName,'EmailAddress':self._GetSenderEmail(msg)}
            MsgMetaData['Recipients']={"RecipientName":msg.To,"CC":msg.CC,'EmailAddress':self._GetRecipientsEmail(msg)}
            MsgMetaData['Subject']=msg.Subject
            extraction['MsgMetaData']=MsgMetaData
            extraction['Msg_Body']=msg.Body
            AttList=[]

            for inx,att in enumerate(msg.Attachments):
                AttFileName = MsgAttachment_directory + "\\" + MsgFileName + "_Attachment" +str(inx+1) + "_" + att.FileName
                if not os.path.exists(MsgAttachment_directory):
                    os.makedirs(MsgAttachment_directory)
                att.SaveAsFile(os.path.abspath(AttFileName))
#                 if is_GetAtt == True:
#                     self._GetAttachmentJson(AttFileName)   
                AttList.append(AttFileName)
            extraction['AttachmentList']= AttList  #Write a function to get a dictionary of Attachment
            extraction['SubmissionNumber']=_GetSubmissionNumber(filename)
            extraction['FileType']=filename.split("\\")[-1].split('.')[-1].lower()
            extraction['FileID']=hashlib.md5(str(extraction).encode('utf-8')).hexdigest()
            extraction['FileLocation']=filename
            elapse_time = time.time() - start_time
            extraction['TotalElapsedTimeMs'] = str(elapse_time/60.0)
            date_time = datetime.datetime.utcnow()
            extraction['ProcessedDateTime'] = str(date_time)
            msg.Close(True)
            del outlook
            return extraction
        except:
            msg.Close(True) if msg is not None else False
            del msg
            del outlook 
            traceback.print_exc()
    def dump2Json(self,filename):  
        JsonFileName = self.CreateJSON._GetJsonFileName(filename)
        if not os.path.exists(JsonFileName):
            extraction = self._ParseOutlookFile(filename)
            if extraction is None:
                raise ValueError("Extraction is None")
            self.CreateJSON.ToJson(filename,extraction)
class WordConverter():
    def __init__(self):
        pass
    def word2pdf(self, filename): ## Word Filename
        import comtypes.client
        try:
            file_format = filename.split('.')[-1].lower()
            if file_format == "doc":
                DocFileName=filename.split('\\')[-1].lower().replace('.doc','_doc.pdf')
            elif file_format == "docx":
                DocFileName=filename.split('\\')[-1].lower().replace('.docx','_docx.pdf')
            DocFileDirectory = '\\'.join(filename.split('\\')[-6:-1])
            word2pdfFolder = SharedFolderInitial+':\\Starr Sprint-F18\\Data\\360-documents-word2pdf\\'
            PdfFileDirectory=os.path.join(word2pdfFolder,DocFileDirectory)
            if not os.path.exists(PdfFileDirectory):
                os.makedirs(PdfFileDirectory)
            PdfFileName=os.path.join(PdfFileDirectory,DocFileName)
            in_file = os.path.abspath(filename)
            out_file = os.path.abspath(PdfFileName)


            doc = None
            word = comtypes.client.CreateObject('Word.Application')
            doc = word.Documents.Open(in_file)
            doc.SaveAs(out_file, FileFormat=17)
            doc.Close()
            del doc
            word.Quit()
            del word
        except:
            doc.Close() if doc is not None else False
            del doc
            word.Quit()
            del word
            traceback.print_exc() 
class PdfParser():
    def __init__(self):
        self.CreateJSON=CreateJSON()
    def _get_new_filename(self,file): ## PNG Filename
        new_filename = file.replace('.pdf','_pdf2pngConversion').replace('.PDF','_pdf2pngConversion')
        return new_filename
    def pdf2png(self):
        from wand.image import Image
        from PyPDF2 import PdfFileReader
        ## Split PDF file into pages
        file=self.filename.split('\\')[-1]
        filepath ='\\'.join(self.filename.split('\\')[:-1])
        FileDirectory = '\\'.join(self.filename.split('\\')[-6:])
        pngFolder = SharedFolderInitial+":\\Starr Sprint-F18\\Data\\360-documents-png\\"
        new_directory=os.path.join(pngFolder,FileDirectory)
        if not os.path.exists(new_directory):
            os.makedirs(new_directory)
        exist_pages = len(os.listdir(new_directory))
        
        pdf_pages = PdfFileReader(open(self.filename,'rb')).getNumPages()
        if exist_pages == pdf_pages:
            return
        with(Image(filename=os.path.abspath(self.filename),resolution=300)) as source:
            images=source.sequence
            pages=len(images)            
            for k in range(pages):
                ## Convert each pdf page to png
                new_png_file = self._get_new_filename(file)+'Page'+str(k+1)+'.png'
                new_png_filepath = os.path.join(new_directory,new_png_file)
                # save png in new and separate folder
                if not os.path.exists(new_png_filepath):
                    Image(images[k]).save(filename=new_png_filepath)

    def _vision_api(self,png_filepath): ## apply computer vision on image and get dictionary format response
        import requests
        try:
            with open(png_filepath, "rb") as image_file:
                files = {'field_name': image_file}
                headers  = {'Ocp-Apim-Subscription-Key': subscription_key}
                params   = {'language': 'unk', 'detectOrientation ': 'true'}
                response = requests.post(ocr_url, headers=headers, params=params, files=files)
                response.raise_for_status()
                analysis = response.json()
            return analysis
        except:
            traceback.print_exc()


    def _OCR_response(self): ## PDF Filename
        
        file=self.filename.split('\\')[-1]
        FileDirectory = '\\'.join(self.filename.split('\\')[-6:])
        pngFolder = SharedFolderInitial+":\\Starr Sprint-F18\\Data\\360-documents-png\\"
        PNG_directory=os.path.join(pngFolder,FileDirectory)

        page_num=0
        for page in os.listdir(PNG_directory):
            if page.split(".")[-1].lower() == "png":
                page_num+=1
        response = []       
        for k in range(page_num):
            png_file = self._get_new_filename(file)+'Page'+str(k+1)+'.png'
            png_filepath = os.path.join(PNG_directory,png_file)

            png_extraction={}
            png_extraction = self._vision_api(png_filepath)
            if png_extraction:
                png_extraction['FileLocation_PNG'] = png_filepath
                response.append(png_extraction)
            else:
                png_extraction={}
                png_extraction['FileLocation_PNG'] = png_filepath
                response.append(png_extraction)
        if response:
            return response
    
    def _concatContent(self, response): # concat text to one string
        contentextracted = []
        for analysis in response:
            for region in analysis['regions']:
                for line in region['lines']:
                    
                    for word in line['words']:

                        contentextracted.append(word['text'])
        contentextracted = ' '.join(contentextracted)
        return contentextracted
    
    def _ParsePdfFile(self): ## create JSON and save to database
        start_time = time.time()
        ## Identify file format for standardization
        file_format = self.filename.split(".")[-1].lower()

        Response = self._OCR_response()

        extraction = {}
        extraction['PDF_Response'] = Response
        content_concated = self._concatContent(Response)
        extraction['Content_Extracted'] = content_concated
        extraction['SubmissionNumber']=_GetSubmissionNumber(self.filename)
        extraction['FileType']=file_format
        extraction['FileID']=hashlib.md5(str(extraction).encode('utf-8')).hexdigest()
        extraction['FileLocation']=self.filename
        elapse_time = time.time() - start_time
        date_time = datetime.datetime.utcnow()
        extraction['TotalElapsedTimeMs'] = str(elapse_time/60.0)
        extraction['ProcessedDateTime'] = str(date_time)
        return extraction
    def dump2Json(self,filename):
        self.filename = filename
        JsonFileName = self.CreateJSON._GetJsonFileName(self.filename)
        if not os.path.exists(JsonFileName):
            extraction = self._ParsePdfFile()
            if extraction is None:
                raise ValueError("Extraction is None")
            self.CreateJSON.ToJson(self.filename,extraction)
class Workflow():
    def __init__(self):
        self.OutlookParser = OutlookParser()
        self.ExcelParser = ExcelParser()
        self.WordConverter=WordConverter()
        self.PdfParser=PdfParser()
    def execute_workflow(self,filename): # identification of input file type
        file_format=filename.split(".")[-1].lower()
        if file_format in ['xls', 'xlsx', 'csv']:
            if is_getExcel==True:
                self.ExcelParser.dump2Json(filename)
        elif file_format in ['doc', 'docx']:
            if is_getWord==True:
                self.WordConverter.word2pdf(filename)
        elif file_format in ["pdf"]:
            if is_pdf2png==True:
                self.PdfParser.pdf2png(filename)
            if is_getPdf==True:
                self.PdfParser.dump2Json(filename)
        elif file_format in ['msg']:
            if is_getMsg==True:
                self.OutlookParser.dump2Json(filename)
if __name__ == "__main__":
    import argparse
    # parse cmd args
    parser = argparse.ArgumentParser(
            description="Data Scraper"
        )
    parser.add_argument('--Ini', action="store", dest="SharedFolderInitial", required=True)
    parser.add_argument('--F', action="store", dest="FileList", required=False)
    parser.add_argument('--Vs', action="store", dest="visionAPI", required=False)
   
    args =  vars(parser.parse_args())
#     print(args)

    # set SharedFolderInitial
    SharedFolderInitial = args['SharedFolderInitial']
    FileList = args['FileList']
    visionAPI = args['visionAPI']
    
    if visionAPI is not None:
        with open("visionAPI.json", "r") as V:   #Pickling
            visionAPI = json.load(V)
        subscription_key=visionAPI['subscription_key']
        vision_base_url=visionAPI['vision_base_url']
        ocr_url = vision_base_url + "/ocr"
    
    if FileList is not None:
        with open(FileList, "rb") as fp:   #Pickling
            file_list=list(pickle.load(fp))
    else:
        file_list = [input("Enter File Path:")]
    NumStart=len(file_list)
    ProcessedFileList=[]
    AttIsMsg=[]
    for eachfile in file_list[:]:
        workflow =Workflow()
        print("Processing " + eachfile)
        try:
            workflow.execute_workflow(filename=eachfile)
            ProcessedFileList.append(eachfile)
        except:
            traceback.print_exc()
            continue
        file_list.remove(eachfile)
        NumLeft = len(file_list)
        FinishPercent = NumLeft / np.float(NumStart)
        print(str(np.round(FinishPercent*100,2)) + "% Completed, " + str(NumLeft) + " Left.")

    with open("Unprocessed File List.pickle", "wb") as fp:   #Pickling
        pickle.dump(file_list, fp)
    with open("Processed File List.pickle", "wb") as fp:   #Pickling
        pickle.dump(ProcessedFileList, fp)

