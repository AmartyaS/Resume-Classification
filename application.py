# -*- coding: utf-8 -*-
"""
Created on Tue Nov 30 05:21:07 2021

@author: ASUS
"""

import os
import re
#import docx
import nltk
import spacy
import pickle
import PyPDF2
import textract
import docx2txt
import jsonify
import requests
import pandas as pd
from venn import venn
from sklearn import metrics
from sklearn.svm import SVC
import matplotlib.pyplot as plt
#import win32com.client as win32
from sklearn import linear_model
from sklearn import preprocessing
from pyresparser import ResumeParser
#from win32com.client import constants
from matplotlib.gridspec import GridSpec
from sklearn.naive_bayes import MultinomialNB
from matplotlib_venn import venn2, venn2_circles
from matplotlib_venn import venn3, venn3_circles
from flask import Flask, render_template, request
from sklearn.metrics import classification_report
from sklearn.neighbors import KNeighborsClassifier
from sklearn.multiclass import OneVsRestClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import TfidfVectorizer
#from __future__ import division, print_function
# coding=utf-8
import sys
import os
import glob
import re
import numpy as np
import pandas as pd

import spacy
import docx2txt
import PyPDF2
from glob import glob
import pickle

# Flask utils
from flask import Flask, redirect, url_for, request, render_template
from werkzeug.utils import secure_filename
from gevent.pywsgi import WSGIServer

# Define a flask app
app = Flask(__name__)


model = pickle.load(open('Resume_Classification.pkl', 'rb'))
vocab = pickle.load(open('Words.pkl', 'rb'))

# Model
def model_predict(path,model):
    file=scan_cv(path)
    dat=skillset(file)
    lines=' '.join([i for i in dat])
    x=vocab.transform([lines])
    prediction=model.predict(x)
    return prediction
    
# Building the data file
def portfolio(raw):
    raw=re.sub(r'\n\s*\n','\n',str(raw),re.MULTILINE)
    raw=raw.replace('\\n',' ')
    raw=raw.replace('\\t','')
    raw=raw.lower()
    return raw
             
def skillset(data):
    data=portfolio(data)
    word=nltk.word_tokenize(data)
    # Combined Skillset repo
    skill=['css', 'xhtml', 'mongodb', 'sql', 'bootstrap', 'java', 'node.js', 
           'html5', 'php', 'react', 'javascript', 'redux(knowledge)', 'react js', 
           'mysql', 'html 4', 'html', 'c', 'css 3', 'sass', 'nodejs', 'aws', 
           'aws/azure', 'html 5', 'xml', 'css3', 'nestjs', 'react.js', 'reactjs', 
           'redux', 'css3/bootstrap', 'xampp','weblogic', 'java', 'ms-sql', 'toad',
           'oracle', 'putty', 'sqr', 'sql', 'fscm', 'hcm', 'tuxedo', 'docker', 
           'peoplesoft', 'ansible','picof', 'xslt', 'peoplesoft', 'ccb', 'xpath', 
           'x-path', 'sql', 'hcm', 'xml', 'eib', 'peci', 'birt', 'workday', 'ccw',
           't-sql', 'msbi', 'mysql', 'ssasssrs', 'aws', 'hive', 'bcp', 'ssms', 
           'query', 'r', 'ssdt', 'teradata', 'putty', 'sql', 'plsql', 'bi', 'ssis',
           'tableau', 'excel', 'python', 'report', 'bids', 'redshift', 'etl']
    skill=list(set(skill))
    matches=list(set(word).intersection(set(skill)))
    html=["html","html 5", "html5","html 4","html4"]
    react=["react js","react","react.js"]
    node=["nodejs","node.js","node js"]
    css=["css","css3","css4","css5"]
    bstp=["css","javascript"]
    picof=["picof","peci"]
    ccw=["ccw","ccb"]
    xslt=["x-path","xpath","xslt","xml"]
    wch=["workday","hcm"]
    ssrs=['report','ssrs']
    sql=['sql','mysql','t-sql','plsql','ssas','ssms']
    bi=['bi','msbi']
    etl=['etl','ssis','query']
    plang=['python','r']
    aws=['aws','redshift']
    bids=['ssdt','bids']
    matches=["sql" if i in sql else i for i in matches]
    matches=["bi" if i in bi else i for i in matches]
    matches=["etl" if i in etl else i for i in matches]
    matches=["ssrs" if i in ssrs else i for i in matches]
    matches=["Programming_Lang" if i in plang else i for i in matches]
    matches=["AWS-Redshift" if i in aws else i for i in matches]
    matches=["bids" if i in bids else i for i in matches]
    matches=["html" if i in html else i for i in matches]
    matches=["reactjs" if i in react else i for i in matches]
    matches=["nodejs" if i in node else i for i in matches]
    matches=["css" if i in css else i for i in matches]
    matches=["picof" if i in picof else i for i in matches]
    matches=["ccw" if i in ccw else i for i in matches]
    matches=["xslt" if i in xslt else i for i in matches]
    matches=["workday-hcm" if i in wch else i for i in matches]
    matches=list(set(matches))
    matches.sort()
    c=0
    for n,i in enumerate(matches):
        if i in bstp:
            c+=1
            if c>=2:
                matches[n]="bootstrap"
                if "css" in matches:
                    matches.remove("css")
                elif "javascript" in matches:
                    matches.remove("javascript")
                else:
                    matches.sort()
        elif i=="css3/bootstrap":
            matches[n]="bootstrap"
        else:
            c+=0
            matches=list(set(matches))
            return matches
        
def scan_cv(resume_path):
    text=[ ]
    if(resume_path.endswith('docx')):
        info=docx2txt.process(resume_path)
        text.append(info)
        return text
    if(resume_path.endswith('pdf')):
        count=0
        pdfreader=PyPDF2.PdfFileReader(open(resume_path,'rb'))
        page_num=pdfreader.getNumPages()
        while count<page_num:
            page_info=pdfreader.getPage(count)
            count+=1
            info=page_info.extractText()
            text.append(info)
            return text
        
@app.route('/', methods=['GET'])
def index():
    # Main page
    return render_template('index.html')

@app.route('/predict', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        # Get the file from post request
        f = request.files["file"]

        # Save the file to ./uploads
        basepath = os.path.dirname(__file__)
        file_path = os.path.join(
            basepath, 'uploads', secure_filename(f.filename))
        f.save(file_path)

        # Make prediction
        preds = model_predict(file_path, model)
        if(preds==1):
            result='Designation is ReactJS Developer'
        elif(preds==0):
            result='Designation is PeopleSoft Developer'
        elif(preds==2):
            result='Designation is WorkDay Developer'
        elif(preds==3):
            result='Designation is SQL Developer'
        else:
            result='Kindly upload the resume in docx or pdf format'


              # Convert to string
        return result
    return None


if __name__ == '__main__':
    app.run(debug=True)