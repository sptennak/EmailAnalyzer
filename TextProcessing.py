# -*- coding: utf-8 -*-
"""
Created on Mon ‎May 21 21:17:45 2018

@author: Sumudu Tennakoon
"""

###############################################################################
#RegEx
EMAIL_FORMAT = re.compile(r'([\w\.-]+@[\w\.-]+\.\w+)')
DOMAIN_FORMAT = re.compile(r'@([\w\.-]+)')
NAME_EMAIL_FORMAT = re.compile(r'\W*(.*?)\W*([\w\.-]+@[\w\.-]+\.\w+)')
MESSAGID_FORMAT = re.compile(r'<(.*?)>')
REFERECES_FORMAT = re.compile(r'[,<](.*?)[,>]') #This need to be revised. in the case of '<ABC>,<DEF>' if will return '<DEF'
###############################################################################
###############################################################################
# Supporting Functions
###############################################################################
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
                '&amp;'   :'&'
                }
    
        for (pattern,sub) in Entity.items():
            Text = Text.replace(pattern, sub)  
    
        Text=re.sub(r'[\t][\t]+' , '\t', Text)
        Text=re.sub(r'[\r][\r]+' , '\r', Text)
        Text=re.sub(r'[\n][\s]*[\n]+' , '\n', Text)
        Text=re.sub(r'^[\s]+' , '', Text)
        Text=re.sub(r'[\n][\s]+$' , '', Text)
    else:
        pass
        
    return Text

###############################################################################
#Data Cleanup
###############################################################################
def GetEmails(EmailColumn):
    return re.findall(EMAIL_FORMAT, str(EmailColumn))

def GetMessageIDs(EmailColumn):
    return re.findall(MESSAGID_FORMAT, str(EmailColumn))

def GetReferenceList(EmailColumn):
    return re.findall(REFERECES_FORMAT, str(EmailColumn))