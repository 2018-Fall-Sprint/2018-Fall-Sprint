
# coding: utf-8


import os
import re
import sys
import time
import shutil
import datetime
import warnings
import requests
import pandas as pd
import numpy as np
import comtypes.client # word2pdf
from wand.image import Image
from PIL import Image as Image2
import matplotlib.pyplot as plt
from win32com.client import Dispatch
import win32com
from matplotlib.patches import Rectangle
import hashlib
import json
import pickle
import traceback

class WorkFlow():
    def __init__(self, filename):
        self.filename = filename   
    def _GetPolicyNumber(self,filename):
        PolicyNumber = filename.split('\\')[-2]
        return PolicyNumber
    def _GetJsonFileName(self,filename,file_format):
        File = "".join(filename.split("\\")[-1].split('.')[:-1])
        JsonFileDirectory=SharedFolderInitial+':\\Starr Sprint-F18\\Data\\JSON Output\\'
        if not os.path.exists(JsonFileDirectory):
            os.makedirs(JsonFileDirectory)
        JsonFileName = File +"_"+ file_format +".json"
        return os.path.join(JsonFileDirectory,JsonFileName)    
    def _ToJson(self,Jsonfilename,extraction): 
        with open(Jsonfilename, 'w') as outfile:
            json.dump(extraction,outfile)
            print(Jsonfilename + ' Processed.')
            print()
    def _word2PDF(self, filename): ## Word Filename
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
            return PdfFileName
        except:
            doc.Close() if doc is not None else False
            del doc
            word.Quit()
            del word
            traceback.print_exc()
        

    def _get_new_filename(self, filename): ## PNG Filename
        new_filename = filename.replace('.pdf','_pdf2pngConversion').replace('.PDF','_pdf2pngConversion')
        return new_filename
    
    
    def _vision_api(self, filename): ## apply computer vision on image and get dictionary format response
        try:
            with open(filename, "rb") as image_file:
                files = {'field_name': image_file}
                headers  = {'Ocp-Apim-Subscription-Key': subscription_key}
                params   = {'language': 'unk', 'detectOrientation ': 'true'}
                response = requests.post(ocr_url, headers=headers, params=params, files=files)
                response.raise_for_status()
                analysis = response.json()
            return analysis
        except:
            traceback.print_exc()
    
    def _vision_api_response(self, filename,timesleep): ## PDF Filename
        
        ## Split PDF file into pages
        file=filename.split('\\')[-1]
        filepath ='\\'.join(filename.split('\\')[:-1])
        FileDirectory = '\\'.join(filename.split('\\')[-6:-1])
        pngFolder = SharedFolderInitial+":\\Starr Sprint-F18\\Data\\360-documents-png\\"
        new_directory=os.path.join(pngFolder,FileDirectory)
        if not os.path.exists(new_directory):
            os.makedirs(new_directory)
        with(Image(filename=os.path.abspath(filename),resolution=400)) as source:
            images=source.sequence
            pages=len(images)
            LastPage_file = self._get_new_filename(file)+'Page'+str(pages)+'.png'
            LastPage_filename = os.path.join(new_directory,LastPage_file)
            response = []
            if not os.path.exists(LastPage_filename):
                for k in range(pages):
                    ## Convert each pdf page to png
                    new_file = self._get_new_filename(file)+'Page'+str(k+1)+'.png'
                    new_filename = os.path.join(new_directory,new_file)
                    # save png in new and separate folder
                    Image(images[k]).save(filename=new_filename)
                    ## Apply VisionAPI
                    png_extraction = self._vision_api(new_filename)
                    ## Add PNG location
                    png_extraction['FileLocation_PNG'] = new_filename
                    ## Append dictionary
                    response.append(png_extraction)
                    time.sleep(timesleep)
            else:
                for k in range(pages):
                    png_extraction = self._vision_api(new_filename)
                    ## Add PNG location
                    png_extraction['FileLocation_PNG'] = new_filename
                    ## Append dictionary
                    response.append(png_extraction)
                    time.sleep(timesleep)
        return response
    
    def _boundingbox_R(self, response): ## for readability
        boundingbox = []
        for analysis in response:
            BBOX = {}
            bbox = {}
            BBOX['FileLocation_PNG'] = analysis['FileLocation_PNG']
            for region in analysis['regions']:
                for line in region['lines']:
                    text = []
                    for word in line['words']:
                        text.append(word['text'])
                    text = ' '.join(text)
                    bbox[line['boundingBox']] = text
            BBOX['Bounding_Box'] = bbox
            boundingbox.append(BBOX)
        return boundingbox  

    def _num_to_col_letters(self, num):
        letters = ''
        while num:
            mod = (num - 1) % 26
            letters += chr(mod + 65)
            num = (num - 1) // 26
        return ''.join(reversed(letters))
    
    def _extract_Excel(self, filename): # Excel Filename, end with xls/xlsx/csv
        try:
            wb=None
            excel = Dispatch('Excel.Application')
            wb = excel.Workbooks.Open(os.path.abspath(filename))
            wb.CheckCompatibility=False
            excel_texts = {}       
            for page in range(excel.Worksheets.Count):
                if excel.Worksheets(page+1).UsedRange() is None:
                    continue
                pagename = excel.Worksheets[page+1].Name
                excel_text = [list(L) for L in excel.Worksheets(page+1).UsedRange()]
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

    def ParsePdfFile(self, filename,timesleep): ## create JSON and save to database
        if self._extract_Excel(filename) is not None:
            start_time = time.time()
            ## Identify file format for standardization
            file_format = filename.split(".")[-1].lower()
            if file_format == "pdf":
                pdf_filename = filename
            elif file_format == "doc" or file_format == "docx":
                pdf_filename =self._word2PDF(filename)
            extraction = {}
            vision_response = self._vision_api_response(pdf_filename,timesleep)
            bounding_box_R = self._boundingbox_R(vision_response)
            extraction['PDF_Response'] = bounding_box_R
            content_extracted_R = self.content_extraction_pdf(bounding_box_R)
            extraction['Content_Extracted'] = content_extracted_R
            extraction['PolicyNumber']=self._GetPolicyNumber(filename)
            extraction['FileType']=filename.split("\\")[-1].split('.')[-1].lower()
            extraction['FileID']=hashlib.md5(str(extraction).encode('utf-8')).hexdigest()
            extraction['FileLocation']=filename
            elapse_time = time.time() - start_time
            date_time = datetime.datetime.utcnow()
            extraction['TotalElapsedTimeMs'] = str(elapse_time/60.0)
            extraction['ProcessedDateTime'] = str(date_time)
            return extraction  
    def ParseExcelFile(self, filename):
        if self._extract_Excel(filename) is not None:
            start_time = time.time()
            excel_texts = self._extract_Excel(filename)
            extraction = {}  
            extraction['Excel_Response'] = excel_texts
            content_extracted = self._content_extraction_excel(excel_texts)
            extraction['Content_Extracted'] = content_extracted
            extraction['PolicyNumber']=self._GetPolicyNumber(filename)
            extraction['FileType']=filename.split("\\")[-1].split('.')[-1].lower()
            extraction['FileID']=hashlib.md5(str(extraction).encode('utf-8')).hexdigest()
            extraction['FileLocation']=filename
            elapse_time = time.time() - start_time
            extraction['TotalElapsedTimeMs'] = str(elapse_time/60.0)  
            date_time = datetime.datetime.utcnow()
            extraction['ProcessedDateTime'] = str(date_time)
            return extraction    
    def _GetAttachmentJson(self,Attachment,timesleep):
        file_format = Attachment.split('.')[-1].lower()
        if file_format in ['xls', 'xlsx', 'csv']:
            JsonFileName = self._GetJsonFileName(Attachment,file_format)
            if not os.path.exists(JsonFileName):
                extraction = self.ParseExcelFile(Attachment)
                self._ToJson(JsonFileName,extraction)
        elif file_format in ['doc', 'docx','pdf']:
            JsonFileName = self._GetJsonFileName(Attachment,file_format)
            if not os.path.exists(JsonFileName):
                extraction = self.ParsePdfFile(Attachment,timesleep)
                self._ToJson(JsonFileName,extraction)
        elif file_format in ['msg']:
            AttIsMsg.append(Attachment)
#             JsonFileName = self._GetJsonFileName(Attachment,file_format)
#             if not os.path.exists(JsonFileName):
#                 extraction = self.ParseOutlookFile(Attachment,timesleep)
#                 self._ToJson(JsonFileName,extraction)
    def ParseOutlookFile(self,filename,timesleep):
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
                self._GetAttachmentJson(AttFileName,timesleep)   
                AttList.append(AttFileName)
            extraction['AttachmentList']= AttList  #Write a function to get a dictionary of Attachment
            extraction['PolicyNumber']=self._GetPolicyNumber(filename)
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
    def execute_workflow(self,filename,timesleep): # identification of input file type
        file_format = filename.split('.')[-1].lower()
        if file_format in ['xls', 'xlsx', 'csv']:
            JsonFileName = self._GetJsonFileName(filename,file_format)
            if not os.path.exists(JsonFileName):
                extraction = self.ParseExcelFile(filename)
                self._ToJson(JsonFileName,extraction)
        elif file_format in ['doc', 'docx','pdf']:
            JsonFileName = self._GetJsonFileName(filename,file_format)
            if not os.path.exists(JsonFileName):
                extraction = self.ParsePdfFile(filename,timesleep)
                self._ToJson(JsonFileName,extraction)
        elif file_format in ['msg']:
            JsonFileName = self._GetJsonFileName(filename,file_format)
            if not os.path.exists(JsonFileName):
                extraction = self.ParseOutlookFile(filename,timesleep)
                self._ToJson(JsonFileName,extraction)
if __name__ == "__main__":
    PickleFileName = input("Please Enter FileList File Name. Such as FileList1.pickle:  ")
    SharedFolderInitial = input("Please Enter Your Shared Folder Initial. Such as Z:   ")
    subscription_key = input("Please Enter Your Vision API Subscription Key:   ")
    vision_base_url = input("Please Enter Your Vision API Base URL:   ")
    subscription_key = '798fad47d608432a9786f910a0b919db'
    vision_base_url = "https://southcentralus.api.cognitive.microsoft.com/vision/v1.0"
    ocr_url = vision_base_url + "/ocr"

    with open(PickleFileName, "rb") as fp:   #Pickling
        file_list=list(pickle.load(fp))

    NumStart=len(file_list)
    ProcessedFileList=[]
    AttIsMsg=[]
    for eachfile in file_list[:]:
        Extraction =WorkFlow(eachfile)
        print("Processing " + eachfile)
        try:
            Extraction.execute_workflow(filename=eachfile,timesleep=0)
            ProcessedFileList.append(eachfile)
        except:
            continue
        file_list.remove(eachfile)
        NumLeft = len(file_list)
        FinishPercent = NumLeft / np.float(NumStart)
        print(str(np.round(FinishPercent*100,2)) + "% Completed, " + str(NumLeft) + " Left.")

    with open("Unprocessed File List.pickle", "wb") as fp:   #Pickling
        pickle.dump(file_list, fp)
    with open("Processed File List.pickle", "wb") as fp:   #Pickling
        pickle.dump(ProcessedFileList, fp)

