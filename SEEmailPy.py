# -*- coding: utf-8 -*-
"""
Created on Mon ‎May 21 21:08:09 2018

@author: Sumudu Tennakoon
"""

# System Utilities
import os
import io
import sys
import gc
import traceback

# Email and Text processing
import email
from email.header import decode_header
import re
import uuid # unique ID

# Data handling and analytics tools
import pandas as pd
import numpy as np
from timeit import default_timer as timer

def FileObject2String(FileBytes, FileName='PDF'):
    print(FileName)
    Text = ''
    return Text
###############################################################################
#Data Cleanup
def RemoveTagsFormats(Text):
    if Text!='':           
        # apply rules in given order!
        rules =  [
                { r'[\x20][\x20]+' : u' '},                   # Remove Consecutive spaces
                { r'\s*<br\s*/?>\s*' : u'\n'},      # Convert <br> to Newline 
                { r'</(p|h\d)\s*>\s*' : u'\n\n'},   # Add double newline after </p>, </div> and <h1/>
                { r'<head>.*<\s*(/head|body)[^>]*>' : u'' },     # Remove everything from <head> to </head>
                { r'<script>.*<\s*/script[^>]*>' : u'' },     # Remove evrything from <script> to </script> (javascipt)
                { r'<style>.*<\s*/style[^>]*>' : u'' },     # Remove evrything from <style> to </style> (stypesheet)
                { r'<[^<]*?/?>' : u'' },            # remove remaining tags
                ]
    
        for rule in rules:
            for (pattern,sub) in rule.items():
                Text  = re.sub(pattern, sub, Text)
                
        #https://www.w3schools.com/charsets/ref_html_entities_4.asp
        #https://docs.python.org/3/library/html.entities.html#html.entities.html5     
        Entity={
                '&lt;'    :'<', 
                '&gt;'    :'>', 
                '&nbsp;'  :' ', 
                '& nbsp;' :' ',
                '&n bsp;' :' ',
                '&nb sp;' :' ',
                '&nbs p;' :' ',
                '&quot;'  :'"',
                'cent;'   :u'\xA2',
                '&pound;' :u'\xA3',
                '&copy;'  :u'\xA9',
                '&reg;'   :u'\xAE',
                '&plusmn;':u'\xB1',
                '&frac14;':u'\xBC',
                '&frac12;':u'\xBD',
                '&frac34;':u'\xBE',
                '&times;' :u'\xD7',
                '&prime;' :u'\x2032',
                '&Prime'  :u'\x2033',
                '&lowast;':u'\x2217',
                '&ne;'    :u'\x2260',
                '&trade;' :u'\x2122',
                '&#8211;' :u'\x8211',
                '&#8217;' :u'\x8217',
                '&amp;'   :'&',
                }
    
        for (pattern,sub) in Entity.items():
            Text = Text.replace(pattern, sub)  
    
        #Text=re.sub('[\x20][\x20]+' , ' ', Text)
        #Text=re.sub('[\n][\n]+' , '\n', Text)
        #Text=re.sub('[\t][\t]+' , '\t', Text)
        Text=re.sub('[\r][\r]+' , '\r', Text)
        Text=re.sub('[\n][\s]*[\n]+' , '\n', Text)
        Text=re.sub('^[\s]+' , '', Text)
        Text=re.sub('[\n][\s]+$' , '', Text)
    else:
        pass
        
    return Text
###############################################################################
#RegEx
EMAIL_FORMAT = re.compile(r'([\w\.-]+@[\w\.-]+\.\w+)')
NAME_EMAIL_FORMAT = re.compile(r'\W*(.*?)\W*([\w\.-]+@[\w\.-]+\.\w+)')
DOMAIN_FORMAT = re.compile(r'@([\w\.-]+)')
MESSAGID_FORMAT = re.compile(r'<(.*?)>')
REFERENCES_FORMAT = re.compile(r'[,<](.*?)[,>]') #This need to be revised. in the case of '<ABC>,<DEF>' if will return '<DEF'
###############################################################################

###############################################################################
def DecodeEmailItem(Text):
    try:
        if Text!=None:
            dc = decode_header(Text)
            ItemValue = dc[0][0]
            Encoding = dc[0][1]
            if Encoding!=None:
                Text = ItemValue.decode(Encoding)
            else:
                pass
        else:
            pass
    except:
        pass
    return Text
###############################################################################
    
###############################################################################
# Function parsing a single email
###############################################################################
def ParseEmail(EmlFilePath, EmailID=None, HeaderFields = None, Attachments=None, WorkingDir = ''):
    ###########################################################################
    # Set Header Feilds
    if HeaderFields == None:
        HeaderFields = ['DATE', 'FROM', 'TO', 'CC', 'BCC', 'SUBJECT', 'CONTENT-LANGUAGE',  'MESSAGE-ID', 
                        'IN-REPLY-TO', 'REFERENCES', 'RETURN-PATH', 'X-MS-HAS-ATTACH', 'X-ORIGINATING-IP']
    else:
        pass
    
    # Set Email ID
    if EmailID == None:
        try:
            EmailID = str(uuid.uuid1()) 
        except:
            pass
    else:
        pass
    ###########################################################################
    # Extract Header    
    HEADER_ITEM = pd.DataFrame(data=[[EmailID, EmlFilePath]], columns=['EmailID', 'FileName'])
    
    try:	
        Email = email.message_from_file(open(EmlFilePath))
        for item in HeaderFields: 
            HEADER_ITEM[item] = DecodeEmailItem(Email.get(item))
    except:
        pass
    ###########################################################################
    # References
    MSGID_ITEMS = pd.DataFrame(data=[], columns=['EmailID', 'Designation', 'MESSAGE-ID'])
    try:
        MID = pd.DataFrame(data=re.findall(MESSAGID_FORMAT, HEADER_ITEM['MESSAGE-ID'].values[0]), columns=['MESSAGE-ID'])
        MID['Designation'] = 0
        MID['Order'] = (MID.index+1).astype('int')
    except:
        pass
    try:
        MID = pd.DataFrame(data=re.findall(MESSAGID_FORMAT, HEADER_ITEM['IN-REPLY-TO'].values[0]), columns=['MESSAGE-ID'])
        MID['Designation'] = 1
        MID['Order'] = (MID.index+1).astype('int')
    except:
        pass
    
    try:
        MID = pd.DataFrame(data=re.findall(REFERENCES_FORMAT, HEADER_ITEM['REFERENCES'].values[0]), columns=['MESSAGE-ID'])
        MID['Designation'] = 2
        MID['Order'] = (MID.index+1).astype('int')
    except:
        pass  
    MSGID_ITEMS['EmailID'] = EmailID
    del(MID)
    ###########################################################################
    # Addresses
    ADDRESS_ITEMS = pd.DataFrame(data=[], columns=['EmailID', 'Designation', 'Order', 'Name', 'Address' ])
    try:
        ADD =  pd.DataFrame(data=re.findall(NAME_EMAIL_FORMAT, HEADER_ITEM['FROM'].values[0]),columns=['Name', 'Address'])
        ADD['Designation']=0
        ADD['Order'] = (ADD.index+1).astype('int')
        ADDRESS_ITEMS = ADDRESS_ITEMS.append(ADD, ignore_index=False, sort=False)
    except:
        pass
    try:    
        ADD =  pd.DataFrame(data=re.findall(NAME_EMAIL_FORMAT, HEADER_ITEM['TO'].values[0]),columns=['Name', 'Address'])
        ADD['Designation']=1    
        ADD['Order'] = (ADD.index+1).astype('int')
        ADDRESS_ITEMS = ADDRESS_ITEMS.append(ADD, ignore_index=False, sort=False)
    except:
        pass
    try:         
        ADD =  pd.DataFrame(data=re.findall(NAME_EMAIL_FORMAT, HEADER_ITEM['CC'].values[0]),columns=['Name', 'Address'])
        ADD['Designation']=2 
        ADD['Order'] = (ADD.index+1).astype('int')
        ADDRESS_ITEMS = ADDRESS_ITEMS.append(ADD, ignore_index=False, sort=False)
    except:
        pass
    try:          
        ADD =  pd.DataFrame(data=re.findall(NAME_EMAIL_FORMAT, HEADER_ITEM['BCC'].values[0]),columns=['Name', 'Address'])
        ADD['Designation']=3 
        ADD['Order'] = (ADD.index+1).astype('int')
        ADDRESS_ITEMS = ADDRESS_ITEMS.append(ADD, ignore_index=False, sort=False)
    except:
        pass  
    ADDRESS_ITEMS['EmailID'] = EmailID
    del(ADD)
    ###########################################################################
    # Content
    ContentID = 0 
    ContentParts = [] #Array of contents
    BodyTextContentParts = []            
    ###########################################################################
    # Attchments
    AttachmentParts = [] #Array of attchments
    ###########################################################################   
#    # Extarct Email Content     
    try:
        for part in Email.walk():   
            ContentMainType = part.get_content_maintype()
            ContentSubType = part.get_content_subtype()
            
            if ContentMainType != 'multipart': #At the loweset level in the content hirachy
                isBodyText=False
                ContentID = ContentID + 1
                ContentString=''
                ContentDisposition = part.get_content_disposition()
                ContentFileName = DecodeEmailItem(part.get_filename())	
                try:
                    ContentTransferEncoding = part.get('Content-Transfer-Encoding').upper()
                except:
                    ContentTransferEncoding = 'QUOTED-PRINTABLE'
                try:
                    MIMEContentID =  re.findall(MESSAGID_FORMAT, part.get('Content-ID'))[0]                    
                except:
                    MIMEContentID = None
                    
                ContentType = '[{}/{}]'.format(ContentMainType, ContentSubType)

                if ContentMainType == 'text' and (ContentSubType == 'plain' or ContentSubType == 'html'):
                    try:
                        CharSet = part.get_content_charset()
                    except:
                        CharSet = 'iso8859_15'
                    try:
                        ContentString = part.get_payload(decode=True).decode(encoding=CharSet)
                        ContentString = RemoveTagsFormats(ContentString) #RemoveHTML(ContentString) 
                        isBodyText = True
                    except:
                        pass   
                #--------------------------------------------------------------
                elif ContentMainType == 'message' and ContentSubType == 'rfc822':
                    try:
                        ContentString = DecodeEmailItem(part.get('SUBJECT'))
                        isBodyText = True
                    except:
                        pass
                    #Record MessageID
                    #Append Subject, Text Body
                    #Seperate from next selection tasks
                    pass                    
                #--------------------------------------------------------------
                elif ContentMainType == 'audio' or ContentMainType == 'video':
                    pass
                elif ContentMainType == 'application' and (ContentSubType == 'zip' or ContentSubType == 'x-7z-compressed' or ContentSubType=='x-rar-compressed'):
                    pass                      
                elif (ContentDisposition == 'attachment' or ContentDisposition == 'inline') or ContentFileName != None:
                    if Attachments != None:
                        if ContentFileName == None: #Add condition not to exclude inline and attachemnt
                            ContentFileName = '<< FileName Blanck >>'
                        else:  
                            AttachmentFileName = 'E{}A{}_{}'.format(EmailID, ContentID, ContentFileName)
                            #print(AttachmentFileName)     
                            try:
                                startTime = timer()
                                AttachmentFileData =  part.get_payload(decode=True) 
                                #--------------------------------------------------------------
                                if Attachments == 'save':
                                    file = open('{}\\{}'.format(WorkingDir,AttachmentFileName), 'wb')     
                                    file.write(AttachmentFileData)
                                    file.close()
                                #--------------------------------------------------------------
                                if Attachments == 'extract':
                                    if ContentTransferEncoding=='BASE64':
                                        pass
                                    else:
                                        AttachmentFileData = AttachmentFileData.encode()
                                    
                                    FileBytes=io.BytesIO(AttachmentFileData)   
                                     
                                    ByteLength = len(AttachmentFileData)/1024 #File size in Kilobytes by byte length                                
                                    if ByteLength<=51200:
                                        AttachmentText = FileObject2String(FileBytes, AttachmentFileName)
                                    else:
                                        AttachmentText = ''
        
                                    ExecuteTime = timer() - startTime
                                    
                                    ContentFileSize=int(FileBytes.seek(0,2)/1024) #File size in Kilobytes
                                    AttachmentParts.append([EmailID, ContentID, ContentFileName, ContentFileSize, ContentType, MIMEContentID, ExecuteTime, AttachmentText]) 
                                    FileBytes.close()  
                                #--------------------------------------------------------------
                            except:
                                pass    
                    else:
                        pass
                else:
                    pass                    
                ContentParts.append([EmailID, ContentID, ContentType, ContentDisposition, ContentFileName, MIMEContentID]) 
                
                if isBodyText == True:
                    BodyTextContentParts.append([EmailID, ContentID, ContentType, ContentString])
                else:
                    pass
        #######################################################################
        #Update Header with content Info summar
        HEADER_ITEM['ContentCount'] = ContentID # Last content ID (start 1, 0 if none)
        HEADER_ITEM['FileReadError'] = 0

        #Create content andd attchemnt                          
        CONTENT_ITEMS = pd.DataFrame(data=ContentParts, columns=['EmailID', 'ContentID', 'ContentType', 'ContentDisposition', 'ContentFileName', 'MIMEContentID'])  
        BODY_TEXT_ITEMS = pd.DataFrame(data=BodyTextContentParts, columns=['EmailID', 'ContentID', 'ContentType', 'ContentString'])  
        ATTACHMENT_ITEMS = pd.DataFrame(data=AttachmentParts, columns=['EmailID', 'ContentID', 'ContentFileName', 'ContentFileSize', 'ContentType', 'MIMEContentID', 'ExecuteTime', 'AttachmentText'])
        #######################################################################  
    except:
        print('ParseEmail Exception Occured: {}'.format(traceback.format_exc()))
        HEADER_ITEM= pd.DataFrame(data=[[EmailID, EmlFilePath, 0, 1]], columns=['EmailID', 'FileName', 'ContentCount', 'FileReadError'])   
        BODY_TEXT_ITEMS = pd.DataFrame()
        CONTENT_ITEMS = pd.DataFrame()
        ATTACHMENT_ITEMS = pd.DataFrame()
        
    del(Email)
    del(ContentParts) 
    del(BodyTextContentParts)
    del(AttachmentParts)           
    gc.collect()
   
    return HEADER_ITEM, ADDRESS_ITEMS, MSGID_ITEMS, CONTENT_ITEMS, BODY_TEXT_ITEMS, ATTACHMENT_ITEMS 

###############################################################################
EmlFilePath = r'C:\Users\spten\Documents\GitHub\EmailAnalyzer\sample3.eml'
EmailID=1
header, address, msgid, content, bodytext, attchment = ParseEmail(EmlFilePath, EmailID)
#header, address, msgid= ParseEmail(EmlFilePath, EmailID)