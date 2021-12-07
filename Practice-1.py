# -*- coding: utf-8 -*-
"""
Created on Wed Nov  3 11:14:12 2021

@author: ASUS
"""

import os
import re
#import docx
import nltk
import spacy
import textract
import PyPDF2
import docx2txt
import pandas as pd
from venn import venn
%matplotlib inline
from sklearn import metrics
from sklearn.svm import SVC
import matplotlib.pyplot as plt
import win32com.client as win32
from sklearn import linear_model
from sklearn import preprocessing
from pyresparser import ResumeParser
from win32com.client import constants
from matplotlib.gridspec import GridSpec
from sklearn.naive_bayes import MultinomialNB
from matplotlib_venn import venn2, venn2_circles
from matplotlib_venn import venn3, venn3_circles
from sklearn.metrics import classification_report
from sklearn.neighbors import KNeighborsClassifier
from sklearn.multiclass import OneVsRestClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split

from sklearn.feature_extraction.text import TfidfVectorizer

# Skillset Repository
global skill
skill=['HTML', 'CSS3', 'XML', 'JavaScript', 'JSON', 'React JS', 'Node.js', 'GitHub', 
       'Agile', 'SCRUM', 'Redux', 'HTML5', 'JIRA', 'Bootstrap', 'HTML5', 'CSS3',
       'Bootstrap','JavaScript', 'DOM', 'ReactJS', 'ReactJS', 'C', 'Reactjs', 
       'JavaScript', 'MySQL', 'HTML', 'CSS', 'ReactJs', 'Redux', 'Bootstrap',  'HTML 4',
       'CSS3', 'Sass', 'JavaScript', 'jQuery',  'Agile', 'Html', 'Sql', 'Reactjs', 
       'Nodejs', 'MERN Stack', 'Mongo Database', 'CSS', 'BOOTSTRAP', 'JAVA', 'JAVASCRIPT',
       'MySQL', 'SPRING BOOT', 'REACT.JS','ANGULAR', 'CSS3/Bootstrap', 'HTML5', 'Node.js', 
       'MICROSERVICES', 'MongoDB', 'AWS/AZURE', 'Html5', 'CSS3', 'JavaScript', 'JQuery', 
       'TypeScript', 'ReactJS', 'BootStrap', 'Angular', 'Redux', 'NodeJS', 'MongoDB', 
       'Github', 'HTML5', 'CSS3', 'JavaScript', 'ReactJs', 'Bootstrap', 'C', 'React', 
       'Bootstrap', 'JavaScript', 'JSON', 'NodeJS', 'HTML', 'JavaScript', 'Bootstrap',
       'ReactJS' ,'Hooks', 'Redux(Knowledge)', 'CSS3', 'SASS', 'HTML', 'CSS3', 'XHTML',
       'HTML5', 'CSS', 'JQuery', 'JavaScript',  'Bootstrap', 'React', 'Angular', 'HTML5',
       'CSS3', 'JavaScript', 'JQuery', 'Json', 'Bootstrap', 'ReactJS', 'Mysql', 'NestJS',
       'AngularJs', 'ReactJs', 'NestJs', 'ReactJs', 'Redux', 'HTML', 'CSS', 'Bootstrap', 
       'jQuery', 'JavaScript', 'HTML', 'CSS', 'BOOTSTRAP','JAVASCRIPT', 'JQUERY', 'PHP',
       'Reactjs' , 'Nodejs', 'Apache', 'XAMPP', 'ReactJS', 'Bootstrap', 'HTML 5', 'CSS', 
       'React', 'Hooks', 'Redux', 'NodeJS', 'MySQL', 'MongoDB', 'HTML', 'CSS', 'JS', 
       'ReactJS', 'AWS', 'ReactJS', 'Redux', 'JavaScript', 'HTML5', 'CSS3', 'ES6', 'Hooks']
skill=set(skill)
skill=list(skill)
skill=[i.lower() for i in skill]  # Converting all the skills to lower
skill=['reactjs', 'node.js', 'c', 'java',
       'html5','aws', 'nodejs', 'css3',
       'sql', 'mysql', 'css3/bootstrap','nestjs',  
       'reactjs', 'html', 'nodejs', 'css', 'bootstrap', 
       'react', 'xhtml', 'react js', 'html', 'aws/azure',
       'html 4', 'xampp', 'javascript', 'sass', 
       'redux(knowledge)', 'php', 'react.js', 'html 5', 
       'mysql', 'mongodb', 'bootstrap', 'redux', 'xml','css 3']
skill=set(skill)
skill=list(skill)

print(skill)

# Locating the Resumes in the directory
record=[]
link="D:\Python Project\EXCELR Project - Document Classification\Resumes\ReactJS Developer"
record=[os.path.join(link,f) for f in os.listdir(link) if os.path.isfile(os.path.join(link,f))]

# Scanning the CVs
def scan_cv(route):
    e=os.path.splitext(route)
    ext=e[1]
    text=[]
    if ext==".pdf":
        count=0
        pdfreader=PyPDF2.PdfFileReader(open(route,'rb'))
        page_num=pdfreader.getNumPages()
        while count<page_num:
            page_info=pdfreader.getPage(count)
            count+=1
            info=page_info.extractText()
            text.append(info)
        return text
    else:
        info=docx2txt.process(route)
        text.append(info)
        return text

# Building the data file
def portfolio(raw):
    raw=re.sub(r'\n\s*\n','\n',str(raw),re.MULTILINE)
    raw=raw.replace('\\n',' ')
    raw=raw.replace('\\t','')
    raw=raw.lower()
    return raw

# Extracting Skills of ReactJS Devops
def skillset(data):
    data=portfolio(data)
    word=nltk.word_tokenize(data)
    matches=list(set(word).intersection(set(skill)))
    html=["html","html 5", "html5","html 4","html4"]
    react=["react js","react","react.js"]
    node=["nodejs","node.js","node js"]
    css=["css","css3","css4","css5"]
    bstp=["css","javascript"]
    matches=["html" if i in html else i for i in matches]
    matches=["reactjs" if i in react else i for i in matches]
    matches=["nodejs" if i in node else i for i in matches]
    matches=["css" if i in css else i for i in matches]
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
        
    '''
    for i in matches:
        if i in html:
            i="html"
        elif i in react:
            i="reactjs"
        elif i in node:
            i="nodejs"
        elif i in css:
            i="css"
        else:
            c=0'''
    return matches


# Converting doc file to docx
def savedoc(path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(record[i])
    doc.Activate ()
    # Rename path with .docx
    new_file_abs = os.path.abspath(record[i])
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
    # Save and Close
    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
    doc.Close(False)
    return new_file_abs        # Returning the link of the saved docx file

# Main Function
final_data=[ ]
i = 0 
name=[]
s=[]
while i < len(record):
    l=os.path.splitext(record[i])
    exten=l[1]
    # 
    if exten==".doc":
        p=savedoc(record[i]) # Calling the function to convert doc to docx
        d=ResumeParser(p).get_extracted_data()
        name.append(d["name"])
        file=scan_cv(p)
        dat=skillset(file)
        lines=' '.join([i for i in dat])
        s.append(lines)
        final_data.append(dat)
    else:
        file= scan_cv(record[i])
        d=ResumeParser(record[i]).get_extracted_data()
        dat = skillset(file)
        lines=' '.join([i for i in dat])
        s.append(lines)
        final_data.append(dat)
        name.append(d["name"])
    i +=1
'''
res=[] # Final List after eliminating the duplicate items
for i in final_data:
    if i not in res:
        res.append(i)
'''
# Creating a DataFrame of candidates and their skills
register=pd.DataFrame(columns=["Name","Domain"])
register["Name"]=name
register["Domain"]="ReactJS Developer"
fd=pd.DataFrame(final_data)
sk=fd.rename(columns={0:"Skill 1",1:"Skill 2",2:"Skill 3",3:"Skill 4",
                      4:"Skill 5",5:"Skill 6",6:"Skill 7",7:"Skill 8",8:"Skill 9",
                      9:"Skill 10",10:"Skill 11"},inplace=False)
sheet1=pd.concat([register,sk],axis=1)
s1=pd.DataFrame(columns=["Domain","Skills"])
s1["Skills"]=s
s1["Domain"]="ReactJS Developer"
#sheet1.to_csv("ReactDev.csv",index=False)

###############################################################################
# Peoplesoft Developer
# The scancv, portfolio and savedoc function will be same as above

# Skillset Repo :
pskill=['PeopleSoft', 'Weblogic', 'Tuxedo', 'Oracle', 'Ansible', 'Docker', 'Java',
        'MS-SQL', 'SQL', 'TOAD', 'Putty', 'SQR', 'FSCM' , 'HCM']
pskill=list(set(pskill))
pskill=[i.lower() for i in pskill ]

# Extracting Skills of Peoplesoft Devops
def peopleskill(data):
    data=portfolio(data)
    word=nltk.word_tokenize(data)
    matches=list(set(word).intersection(set(pskill)))
    matches=list(set(matches))
    return matches


# Link of the Peoplesoft Developer Directory
record=[]
link="D:\Python Project\EXCELR Project - Document Classification\Resumes\Peoplesoft resumes\\"
record=[os.path.join(link,f) for f in os.listdir(link) if os.path.isfile(os.path.join(link,f))]

# Main function
corps=[]
i=0
names=[]
s=[]
while i<len(record):
    l=os.path.splitext(record[i])
    exten=l[1]
    if exten==".doc":
        p=savedoc(record[i])
        d=ResumeParser(p).get_extracted_data()
        names.append(d["names"])
        file=scan_cv(p)
        dat=peopleskill(file)
        lines=' '.join([i for i in dat])
        s.append(lines)
        corps.append(dat)
    else:
        d=ResumeParser(record[i]).get_extracted_data()
        names.append(d["name"])
        file=scan_cv(record[i])
        dat=peopleskill(file)
        lines=' '.join([i for i in dat])
        s.append(lines)
        corps.append(dat)
    i +=1
        
# Creating a DataFrame of candidates and their skills
register=pd.DataFrame(columns=["Name","Domain"])
register["Name"]=names
register["Domain"]="PeopleSoft Developer"
fd=pd.DataFrame(corps)
sk=fd.rename(columns={0:"Skill 1",1:"Skill 2",2:"Skill 3",3:"Skill 4",
                      4:"Skill 5",5:"Skill 6",6:"Skill 7",7:"Skill 8",8:"Skill 9",
                      9:"Skill 10",10:"Skill 11"},inplace=False)
sheet2=pd.concat([register,sk],axis=1)
s2=pd.DataFrame(columns=["Domain","Skills"])
s2["Skills"]=s
s2["Domain"]="PeopleSoft Developer"

#sheet2.to_csv("PeoSoftDev.csv",index=False)

###############################################################################
# WorkDay Developer
# The scancv, portfolio and savedoc function will be same as above

# Link of the WorkDay Resumes Directory
record=[]
link="D:\Python Project\EXCELR Project - Document Classification\Resumes\workday resumes\\"
record=[os.path.join(link,f) for f in os.listdir(link) if os.path.isfile(os.path.join(link,f))]

# Workday Skillset Repo
wskill=['Workday', 'HCM', 'EIB', 'PICOF', 'PECI', 'BIRT', 'XML', 
        'Peoplesoft', 'XSLT', 'EIB','CCW','CCB', 'XPATH', 'SQL', 'X-PATH']
wskill=list(set(wskill))
wskill=[i.lower() for i in wskill ]

# Extracting Skills of WorkDay Devops
def workday(data):
    data=portfolio(data)
    word=nltk.word_tokenize(data)
    matches=list(set(word).intersection(set(wskill)))
    picof=["picof","peci"]
    ccw=["ccw","ccb"]
    xslt=["x-path","xpath","xslt","xml"]
    wch=["workday","hcm"]
    matches=["picof" if i in picof else i for i in matches]
    matches=["ccw" if i in ccw else i for i in matches]
    matches=["xslt" if i in xslt else i for i in matches]
    matches=["workday-hcm" if i in wch else i for i in matches]
    matches=list(set(matches))
    return matches

# Main function
corps=[]
i=0
names=[]
s=[]
while i<len(record):
    l=os.path.splitext(record[i])
    exten=l[1]
    if exten==".doc":
        p=savedoc(record[i])
        d=ResumeParser(p).get_extracted_data()
        names.append(d["names"])
        file=scan_cv(p)
        dat=workday(file)
        lines=' '.join([i for i in dat])
        s.append(lines)
        corps.append(dat)
    else:
        d=ResumeParser(record[i]).get_extracted_data()
        names.append(d["name"])
        file=scan_cv(record[i])
        dat=workday(file)
        lines=' '.join([i for i in dat])
        s.append(lines)
        corps.append(dat)
    i +=1

# Creating a DataFrame of candidates and their skills
register=pd.DataFrame(columns=["Name","Domain"])
register["Name"]=names
register["Domain"]="Workday Developer"
fd=pd.DataFrame(corps)
sk=fd.rename(columns={0:"Skill 1",1:"Skill 2",2:"Skill 3",3:"Skill 4",
                      4:"Skill 5",5:"Skill 6",6:"Skill 7",7:"Skill 8",8:"Skill 9",
                      9:"Skill 10",10:"Skill 11"},inplace=False)
sheet3=pd.concat([register,sk],axis=1)
s3=pd.DataFrame(columns=["Domain","Skills"])
s3["Skills"]=s
s3["Domain"]="Workday Developer"


#sheet3.to_csv("WorDayDev.csv",index=False)

###############################################################################
# SQL Developer
# The scancv, portfolio and savedoc function will be same as above

# Link of the WorkDay Resumes Directory
record=[]
link="D:\Python Project\EXCELR Project - Document Classification\Resumes\SQL Developer Lightning insight\\"
record=[os.path.join(link,f) for f in os.listdir(link) if os.path.isfile(os.path.join(link,f))]

sqskill=['SQL', 'T-SQL', 'Python', 'Redshift', 'AWS', 'Teradata', 'SSIS', 'BI', 'MSBI', 
         'Query', 'Report', 'MySQL', 'Tableau', 'R', 'Excel', 'PLSQL', 'Putty', 'Hive','SSAS'
         'SSRS', 'ETL', 'BIDS', 'SSDT', 'BCP', 'SSMS']
sqskill=list(set(sqskill))
sqskill=[i.lower() for i in sqskill ]

# Extracting Skills of WorkDay Devops
def sql(data):
    data=portfolio(data)
    word=nltk.word_tokenize(data)
    matches=list(set(word).intersection(set(sqskill)))
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
    matches=list(set(matches))
    return matches


# Main function
corps=[]
i=0
names=[]
s=[]
while i<len(record):
    l=os.path.splitext(record[i])
    exten=l[1]
    if exten==".doc":
        p=savedoc(record[i])
        d=ResumeParser(p).get_extracted_data()
        names.append(d["names"])
        file=scan_cv(p)
        dat=sql(file)
        lines=' '.join([i for i in dat])
        s.append(lines)
        corps.append(dat)
    else:
        d=ResumeParser(record[i]).get_extracted_data()
        names.append(d["name"])
        file=scan_cv(record[i])
        dat=sql(file)
        lines=' '.join([i for i in dat])
        s.append(lines)
        corps.append(dat)
    i +=1

# Creating a DataFrame of candidates and their skills
register=pd.DataFrame(columns=["Name","Domain"])
register["Name"]=names
register["Domain"]="SQL Developer"
fd=pd.DataFrame(corps)
sk=fd.rename(columns={0:"Skill 1",1:"Skill 2",2:"Skill 3",3:"Skill 4",
                      4:"Skill 5",5:"Skill 6",6:"Skill 7",7:"Skill 8",8:"Skill 9",
                      9:"Skill 10",10:"Skill 11"},inplace=False)
sheet4=pd.concat([register,sk],axis=1)
#sheet4.to_csv("SQLDev.csv",index=False)
s4=pd.DataFrame(columns=["Domain","Skills"])
s4["Domain"]="SQL Developer"
s4["Skills"]=s

##############################################################################

# Concatenating the entire dataset
final_sheet=pd.concat([sheet1,sheet2,sheet3,sheet4],ignore_index=True)
fin_s=pd.concat([s1,s2,s3,s4],ignore_index=True)
final_sheet.to_csv("Clubbed.csv",index=False)

#############################################################################

# Pie-chart distribution
targetCounts = final_sheet['Domain'].value_counts()
targetLabels  = final_sheet['Domain'].unique()
#label=b["Labels"].to_numpy()
# Make square figures and axes
plt.figure(1, figsize=(22,22))
the_grid = GridSpec(2, 2)
cmap = plt.get_cmap('coolwarm')
plt.subplot(the_grid[0, 1], aspect=1, title='DOMAIN DISTRIBUTION')
source_pie = plt.pie(targetCounts, labels=targetLabels, autopct='%1.1f%%', shadow=True)
plt.show()

##############################################################################

# Visualisation of common skills
skill_repo={
    "React_Skill":set(skill),
    "Peoplesoft_Skill":set(pskill),
    "Workday_Skill":set(wskill),
    "SQL_Skill":set(sqskill)}
venn(skill_repo)
plt.title("Common Skills between four positions")
##############################################################################

# Label Encoding the target column

var=["Domain"]
l_encoder=preprocessing.LabelEncoder()
for i in var:
    fin_s[i]=l_encoder.fit_transform(fin_s[i])
    
a=fin_s["Domain"].value_counts(dropna=True, sort=True)
b=pd.DataFrame(final_sheet["Domain"].value_counts(dropna=True, sort=True))
b["Labels"]=a.index
a=b.pop("Labels")
b.insert(0,'Labels',a)
b=b.rename(columns={'Domain':'Counts'})
print(b)
b.index.value_counts()
b["index"]
fin_s.to_csv("Check.csv",index=False)

###############################################################################

fin_s.columns    # Final Dataset for model making
# Dividing the dataset into target and independent columns
req_val=fin_s["Skills"].values
y=fin_s['Domain']
# TfidfVectorizer 
words=TfidfVectorizer(sublinear_tf=True,
                      stop_words='english',
                      max_features=1500)
words.fit(req_val)
x=words.transform(req_val)

# Splitting the Dataset into training and testing data
x_train,x_test,y_train,y_test=train_test_split(x, y, random_state=33, test_size=0.2)
y_train.value_counts()
y_test.value_counts()

##############################################################################

##      MODEL Building      ##
# Model 1 : KNeighborsClassifier
mod1=OneVsRestClassifier(KNeighborsClassifier())
mod1.fit(x_train,y_train)
pred1=mod1.predict(x_test)
tr_pred1=mod1.predict(x_train)
pd.crosstab(y_test,pred1)
acc1=metrics.accuracy_score(y_test,pred1)*100 
tr_acc1=metrics.accuracy_score(y_train,tr_pred1)*100
metrics.confusion_matrix(y_test,pred1)
print("Classification Report : \n",classification_report(y_test,pred1))
print("Accuracy of KNeighbors Classifier on testing dataset :",acc1)
print("Accuracy of KNeighbors Classifier on training dataset :",tr_acc1)


# Model 2 : Support Vector Classifier
mod2=OneVsRestClassifier(SVC())
mod2.fit(x_train,y_train)
pred2=mod2.predict(x_test)
tr_pred2=mod2.predict(x_train)
pd.crosstab(y_test,pred2)
acc2=metrics.accuracy_score(y_test,pred2)*100 
tr_acc2=metrics.accuracy_score(y_train,tr_pred2)*100
metrics.confusion_matrix(y_test,pred2)
print("Classification Report : \n",classification_report(y_test,pred2))
print("Accuracy of Support Vector Classifier on testing dataset :",acc2)
print("Accuracy of Support Vector Classifier on training dataset :",tr_acc2)

# Model 3 : Logistic Regression
mod3=linear_model.LogisticRegression(multi_class='ovr')
mod3.fit(x_train,y_train)
pred3=mod3.predict(x_test)
tr_pred3=mod3.predict(x_train)
pd.crosstab(y_test,pred3)
acc3=metrics.accuracy_score(y_test,pred3)*100 
tr_acc3=metrics.accuracy_score(y_train,tr_pred3)*100
metrics.confusion_matrix(y_test,pred2)
print("Classification Report : \n",classification_report(y_test,pred3))
print("Accuracy of Logistic Regression on testing dataset :",acc3)
print("Accuracy of Logistic Regression on training dataset :",tr_acc3)

# Model 4 : Multinomial Naive-Bayes Classifier
mod4=OneVsRestClassifier(MultinomialNB())
mod4.fit(x_train,y_train)
pred4=mod4.predict(x_test)
tr_pred4=mod4.predict(x_train)
pd.crosstab(y_test,pred4)
acc4=metrics.accuracy_score(y_test,pred4)*100 
tr_acc4=metrics.accuracy_score(y_train,tr_pred4)*100
metrics.confusion_matrix(y_test,pred4)
print("Classification Report : \n",classification_report(y_test,pred4))
print("Accuracy of Multinomial Naive-Bayes Classifier on testing dataset :",acc4)
print("Accuracy of Multinomial Naive-Bayes Classifier on training dataset :",tr_acc4)

# Model 5 : Random Forest Classifier
mod5=RandomForestClassifier()
mod5.fit(x_train,y_train)
pred5=mod5.predict(x_test)
tr_pred5=mod5.predict(x_train)
pd.crosstab(y_test,pred5)
acc5=metrics.accuracy_score(y_test,pred5)*100 
tr_acc5=metrics.accuracy_score(y_train,tr_pred5)*100
metrics.confusion_matrix(y_test,pred5)
print("Classification Report : \n",classification_report(y_test,pred5))
print("Accuracy of Random Forest Classifier on testing dataset :",acc5)
print("Accuracy of Random Forest Classifier on training dataset :",tr_acc5)

###### All Model Report
report=pd.DataFrame(columns=["Model Name","Training Accuracy", "Testing Accuracy"])
report["Model Name"]=["KNeighbors Classifier", "Support Vector Classifier",
                      "Logistic Regression","Multinomial Naive-Bayes Classifier",
                      "Random Forest Classifier"]
report["Training Accuracy"]=[tr_acc1,tr_acc2,tr_acc3,tr_acc4,tr_acc5]
report["Testing Accuracy"]=[acc1,acc2,acc3,acc4,acc5]
print(report)

##############################################################################

