import os
from spire.doc import *
from spire.doc.common import *
import pandas as pd
from pandas import DataFrame
from lxml import etree
import zipfile

import sys, io

buffer = io.StringIO()
sys.stdout = sys.stderr = buffer
ooXMLns = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def comment_script(file: io.FileIO) -> None:
    document = Document()
    # загружаем документ
    document.LoadFromFile(file)
    authors = []
    texts = []
    answers = []
    for i in range(document.Comments.Count):
        comment = document.Comments[i]
        comment_text = ""
        answer_text = ""
        for j in range(comment.Body.Paragraphs.Count):
            paragraph = comment.Body.Paragraphs[j]
            comment_text += paragraph.Text
            comment_author = comment.Format.Author
            if comment_author == "Замараев Вячеслав Викторович":
                answer_text = comment_text
                comment_text = ""
        authors.append(comment_author)
        texts.append(comment_text)
        answers.append(answer_text)
    
    dict_for_pd = {
        'Имя': authors,
        'Текст': texts,
        'Ответ': answers
    }
    df = DataFrame(dict_for_pd)
    df.to_excel('comments.xlsx', columns=['Имя', 'Текст', 'Ответ'], sheet_name='Sheet1', index=False)
    os.startfile("comments.xlsx")

def comment_script_xml(docxFileName):
    docxZip = zipfile.ZipFile(docxFileName)
    commentsXML = docxZip.read('word/comments.xml')
    et = etree.XML(commentsXML)
    comments = et.xpath('//w:comment', namespaces=ooXMLns)
    authors = []
    dates = []
    values = []
    answers = []
    for c in comments:
        authors.append(c.xpath('@w:author', namespaces=ooXMLns)[0])
        dates.append(str(c.xpath('@w:date', namespaces=ooXMLns)[0])[:-10])
        values.append(c.xpath('string(.)', namespaces=ooXMLns))
        answers.append(str(c.xpath('@w:answer', namespaces=ooXMLns))[1:-1])

    dict_for_pd = {
        'Имя': authors,
        'Текст': values,
        'Дата': dates,
        'Ответ': answers
    }

    df = DataFrame(dict_for_pd)
    df.to_excel('comments.xlsx', columns=['Имя', 'Текст', 'Дата', 'Ответ'], sheet_name='Sheet1', index=False)
    os.startfile("comments.xlsx")
