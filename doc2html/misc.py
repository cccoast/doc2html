#coding:utf-8

from win32com.client import Dispatch as handler
from bs4 import BeautifulSoup

word = handler(('Word.Application'))
word.Visible = False
word.DisplayAlerts = False

def _doc2html(doc_path,html_path):
    doc = word.Documents.Open(doc_path)
    doc.SaveAs(html_path,10)
    doc.Close()
    
def _html2doc(html_path,doc_path):
    doc = word.Documents.Open(html_path)
    if doc_path.endswith('docx'):
        doc.SaveAs(doc_path,12)
    else:
        doc.SaveAs(doc_path,0)
        
def str2tag(s):
    return BeautifulSoup(s,'lxml')

def insert_tag(parent,tag_name,content):
    new_tag = str2tag('<' + tag_name + '>' + '</' + tag_name + '>').find(tag_name)
    new_tag.string = content.decode('cp936') if not isinstance(content, str) else content 
    parent.insert(0,new_tag)