#AICER new content log, loops through all the AICER reports and builds a new log based on content created in the last week. Also populates mini summary column by scraping the masterstore
#Developed by Daniel Hutchings

import csv
import pandas as pd
import glob
import os
import fnmatch
import re
import time
import datetime
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
from bs4 import BeautifulSoup, SoupStrainer

import xml.etree.ElementTree as ET
from lxml import etree

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

def dataframefilter(df, search, weekago, type):
    df = df[df.DateFirstPublished.notnull()]
    df['DateFirstPublished'] = pd.to_datetime(df['DateFirstPublished'], dayfirst=False) #Date is in American format hence dayfirst false
    print("Week ago was: " + str(weekago))
    df = df[df.DateFirstPublished.dt.date >= weekago]
    print("\nBuilding new report...")
    #dropping unnecessary columns
    del df['DisplayId'], df['LexisSmartId'], df['OriginalContentItemId'], df['OriginalContentItemPA'], df['PageType'], df['TopicTreeLevel4'], df['TopicTreeLevel5'], df['TopicTreeLevel6'], df['TopicTreeLevel7'], df['TopicTreeLevel8'], df['TopicTreeLevel9'], df['TopicTreeLevel10'], df['VersionFilename'], df['Filename_Or_Address'], df['CreateDate'], df['MajorUpdateFirstPublished'], df['LastPublishedDate'], df['OriginalLastPublishedDate'], df['LastMajorDate'], df['LastMinorDate'], df['LastReviewedDate'], df['LastUnderReviewDate'], df['SupportsMiniSummary']
    df = df.rename(columns={'id': 'Doc ID', 'ContentItemType': 'Content Type', 'TopicTreeLevel1': 'PA', 'TopicTreeLevel2': 'Subtopic', 'Label': 'Doc Title'})    
    if type != 'news': df['Subtopic'] = df['Subtopic'] + ' > ' + df['TopicTreeLevel3'] 
    del df['TopicTreeLevel3']
    
    df['Link'] = search + df['Doc Title']
    return df

def Filter(filename, files, directory, ReportDir, outputfile, outputfileqandas, outputfilenews, weekago):
    print("\nFiltering...")
    search = 'https://www.lexisnexis.com/uk/lexispsl/tax/search?pa=arbitration%2Cbankingandfinance%2Ccommercial%2Ccompetition%2Cconstruction%2Ccorporate%2Ccorporatecrime%2Cdisputeresolution%2Cemployment%2Cenergy%2Cenvironment%2Cfamily%2Cfinancialservices%2Cimmigration%2Cinformationlaw%2Cinhouseadvisor%2Cinsuranceandreinsurance%2Cip%2Clifesciences%2Clocalgovernment%2Cpensions%2Cpersonalinjury%2Cplanning%2Cpracticecompliance%2Cprivateclient%2Cproperty%2Cpropertydisputes%2Cpubliclaw%2Crestructuringandinsolvency%2Criskandcompliance%2Ctax%2Ctmt%2Cwillsandprobate&submitOnce=true&wa_origin=paHomePage&wordwheelFaces=daAjax%2Fwordwheel.faces&query='
        
    df = pd.read_csv(directory + filename, encoding='UTF-8', low_memory=False) #Load csv file into dataframe
    
    df = df[df.TopicTreeLevel1.isin(['Employment', 'Personal Injury', 'Dispute Resolution', 'Family', 'Property', 'Commercial', 'Information Law', 'Planning', 'Property Disputes', 'IP', 'Construction', 'Local Government', 'TMT', 'Arbitration', 'Wills and Probate', 'Private Client', 'Tax', 'In-House Advisor', 'Corporate', 'Restructuring and Insolvency', 'Environment', 'Practice Compliance', 'Public Law', 'Corporate Crime', 'Financial Services', 'Insurance', 'Energy', 'Pensions', 'Banking and Finance', 'Immigration', 'Competition', 'News Analysis', 'Life Sciences and Pharmaceuticals', 'Practice Management', 'Share Schemes', 'Risk and Compliance'])] # These are the relevent PAs we want to keep in the report
    df2 = df[df.ContentItemType.isin(['QandAs'])] # These are the content types to keep in the report
    df3 = df[df.ContentItemType.isin(['NewsAnalysis'])] # These are the content types to keep in the report
    df = df[df.ContentItemType.isin(['Precedent', 'PracticeNote', 'Checklist', 'AtAGlance'])] # These are the content types to keep in the report 
    df2 = dataframefilter(df2, search, weekago, 'qas')
    df3 = dataframefilter(df3, search, weekago, 'news')    
    df = dataframefilter(df, search, weekago, '')

    #sort by defined list
    sorter = ["PracticeNote", "AtAGlance", "Checklist", "Precedent"]
    df['Content Type'] = df['Content Type'].astype("category") # Convert 'Content Type'-column to category and in set the sorter as categories hierarchy
    df['Content Type'].cat.set_categories(sorter, inplace=True)
    df = df.sort_values(['Content Type'])# 'sort' changed to 'sort_values'

    df.to_csv(ReportDir + outputfile, sep=',',index=False, encoding='UTF-8')
    df2.to_csv(ReportDir + outputfileqandas, sep=',',index=False, encoding='UTF-8')
    df3.to_csv(ReportDir + outputfilenews, sep=',',index=False, encoding='UTF-8')
    print('Exported to ' + ReportDir + outputfile)
    print('Exported to ' + outputfileqandas)
    print('Exported to ' + outputfilenews)


def MiniSummary(ReportDir, outputfile, lookupdpsi):
    df = pd.read_csv(ReportDir + outputfile, encoding='UTF-8') #Load csv file into dataframe
    dfdpsi = pd.read_csv(lookupdpsi + 'lookup-dpsis.csv')
    df1 = pd.DataFrame()
    
    print('Searching for minisummary for: ' + outputfile)

    #search for minisummary
    i = 0
    for index, row in df.iterrows():
        
        DocID = str(df.iloc[i,0])
        ContentType = str(df.iloc[i,1])
        PA = str(df.iloc[i,2])
        Subtopic = str(df.iloc[i,3])
        DocTitle = str(df.iloc[i,4])
        DateFirstPublished = str(df.iloc[i,5])
        Link = str(df.iloc[i,7])
        print(PA)
        print(DocID)

        wasChecklist = 'no' #set default to no
        if ContentType == 'Checklist': 
            ContentType = 'PracticeNote' #Checklists are in the practice notes dpsis
            wasChecklist = 'yes'
        if 'News' in outputfile:
            if PA == 'News Analysis': 
                PA = Subtopic            
            filename = '\\\\lngoxfclup24va\\glpfab4\\Build\\0S4D\\Data_RX\\NEWSANALYSIS_AN_' + DocID + '.xml'
            print(str(i) + ': ' + filename)
            try: 
                soup = BeautifulSoup(open(filename),'lxml') 
                try: summary = soup.find('kh:mini-summary').text
                except: summary = 'Not present'
            except: summary = 'Not present'
            list1 = [[DocTitle, summary, ContentType, PA, Subtopic, DocID, DateFirstPublished, Link]]
            df1 = df1.append(list1)

        else: 
            docloc = dfdpsi.loc[(dfdpsi['ContentType'] == ContentType) & (dfdpsi['PA'] == PA), 'path'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of path
            if wasChecklist == 'yes': #revert back to checklist before exporting
                ContentType = 'Checklist'

            dirjoined = os.path.join(docloc, '*.xml') 
            files = glob.iglob(dirjoined) #search directory and add all files to dict
        
            for filename in files:
                if DocID in filename: 
                    print(str(i) + ': ' + filename)
                    soup = BeautifulSoup(open(filename, encoding='utf8'),'lxml') 
                    try: 
                        summary = soup.find('kh:mini-summary').text
                    except: 
                        summary = 'Not present'

                    list1 = [[DocTitle, summary, ContentType, PA, Subtopic, DocID, DateFirstPublished, Link]]
                    df1 = df1.append(list1)
        i=i+1
    df1.to_csv(ReportDir + outputfile, encoding='UTF-8', sep=',',index=False,header=["Doc Title", "Summary", "Content Type", "PA", "Subtopic", "Doc ID", "DateFirstPublished", "Link"])
    
def QandAsOverviewLog(ReportDir, outputfile, OVFilename):
    AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
    df = pd.read_csv(ReportDir + outputfile, encoding='UTF-8') #Load csv file into dataframe
    list1 = [[]]
    list2 = []
    dfOv = pd.DataFrame()

    for PA in AllPAs:
        PAtotal = len(df[df['PA'] == PA])
        list1 = [[PA, PAtotal]]
        dfOv = dfOv.append(list1)
        list2.append(PAtotal)
    dfOv.to_csv(ReportDir + "\\" + OVFilename, sep=',',index=False, header=["PA", "Total number of new QandAs"]) #Output to CSV
    print("QandAs overview exported to: " + ReportDir + "\\" + OVFilename)
    return list2

    
def NewsOverviewLog(ReportDir, outputfile, OVFilename):   
    AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate', 'Accountancy', 'LexisLibrary', 'News', 'Brexit SI sifting alerts', 'End of year reviews', 'Brexit newsletter']    
    AllPAsNews = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'InHouse Advisor', 'Insurance & Reinsurance', 'IP and IT', 'Life Sciences', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk & Compliance', 'Share Incentives', 'Tax', 'TMT', 'Wills and Probate', 'Accountancy', 'LexisLibrary', 'News', 'Brexit SI sifting alerts', 'End of year reviews', 'Brexit newsletter']    
    i = 0
    df = pd.read_csv(ReportDir + outputfile, encoding='UTF-8') #Load csv file into dataframe
    list1 = [[]]
    dfOv = pd.DataFrame()

    for PA in AllPAsNews:
        PAtotal = len(df[df['PA'] == PA])
        list1 = [[AllPAs[i], PAtotal]]
        print(list1)
        dfOv = dfOv.append(list1)
        i=i+1
    dfOv.to_csv(ReportDir + "\\" + OVFilename, sep=',',index=False, header=["PA", "Total number of new News items"]) #Output to CSV
    print("News overview exported to: " + ReportDir + "\\" + OVFilename)

def OverviewLog(ReportDir, outputfile, OVFilename):
    print("Building static overview...")

    dfOv = pd.DataFrame()
    #Adding extra column to the csv file to hold the categories
    df = pd.read_csv(ReportDir + outputfile, encoding='UTF-8') #Load csv file into dataframe
    
    AllPAs = ['Arbitration', 'Banking and Finance', 'Commercial', 'Competition', 'Construction', 'Corporate', 'Corporate Crime', 'Dispute Resolution', 'Employment', 'Energy', 'Environment', 'Family', 'Financial Services', 'Immigration', 'Information Law', 'In-House Advisor', 'Insurance', 'IP', 'Life Sciences and Pharmaceuticals', 'Local Government', 'Pensions', 'Personal Injury', 'Planning', 'Practice Compliance', 'Practice Management', 'Private Client', 'Property', 'Property Disputes', 'Public Law', 'Restructuring and Insolvency', 'Risk and Compliance', 'Share Schemes', 'Tax', 'TMT', 'Wills and Probate']    
    AllPAsDir = ["ARBITRATION", "BANKINGANDFINANCE", "COMMERCIAL", "COMPETITION", "CONSTRUCTION", "CORPORATE", "CORPORATECRIME", "DISPUTERESOLUTION", "EMPLOYMENT", "ENERGY", "ENVIRONMENT", "FAMILYLAW", "FINANCIALSERVICES", "IMMIGRATION", "INFORMATIONLAW", "INHOUSE", "INSURANCEANDREINSURANCE", "IPANDIT", "LIFESCIENCES", "LOCALGOVERNMENT", "PENSIONS", "PERSONALINJURY", "PLANNING", "PRACTICECOMPLIANCE", "PRACTICEMANAGEMENT", "PRIVATECLIENT", "PROPERTY", "PROPERTYDISPUTES", "PUBLICLAW", "RESTRUCTURINGANDINSOLVENCY", "RISKANDCOMPLIANCE", "SHARESCHEMES", "TAXLAW", "TMT", "WILLSANDPROBATE"]
    i = 0

    for PA in AllPAs:        
        try: PAtotal = len(df[df['PA'] == PA])
        except: PAtotal = 0
        try: PNtotal = len(df[(df['Content Type'] == 'PracticeNote') & (df['PA'] == PA)])
        except: PNtotal = 0
        try: OVtotal = len(df[(df['Content Type'] == 'AtAGlance') & (df['PA'] == PA)])
        except: OVtotal = 0
        try: CLtotal = len(df[(df['Content Type'] == 'Checklist') & (df['PA'] == PA)])
        except: CLtotal = 0
        try: PRtotal = len(df[(df['Content Type'] == 'Precedent') & (df['PA'] == PA)])
        except: PRtotal = 0
        list1 = [[PA, PAtotal, PNtotal, OVtotal, CLtotal, PRtotal]]
        print(list1)
        dfOv = dfOv.append(list1)

        PA = AllPAsDir[i] #swap the PA name for the capitalised version in a different list
        i = i+1        
    
    dfOv.to_csv(ReportDir + "\\" + OVFilename, sep=',',index=False, header=["PA", "Total number of new docs", "Practice Notes", "Overviews", "Checklists", "Precedents"]) #Output to CSV
    print("Overview exported to: " + ReportDir + "\\" + OVFilename)
    
def StackedBar(ReportDir, OVFilename, directory, daterange):
    print('Generating bar chart for new content items...')
    dfOv = pd.read_csv(ReportDir + "\\" + OVFilename) #Load csv file into dataframe
    dfOv = dfOv.sort_values(['Total number of new docs'], ascending = False)
    dfOv = dfOv[dfOv['Total number of new docs'] != 0]
    values = dfOv['Total number of new docs']
    objects =  dfOv['PA'] 

    x_pos = np.arange(len(objects))

    dfOv['Practice Notes'].plot(kind='bar', facecolor='#FF8AC4')
    dfOv['Checklists'].plot(kind='bar', bottom=dfOv['Practice Notes'], facecolor='#ACFFBA')
    dfOv['Overviews'].plot(kind='bar', bottom=[i+j for i,j in zip(dfOv['Practice Notes'], dfOv['Checklists'])], facecolor='#91A8E8')
    dfOv['Precedents'].plot(kind='bar', bottom=[i+j+k for i,j,k in zip(dfOv['Practice Notes'], dfOv['Checklists'], dfOv['Overviews'])], facecolor='#E8E2CE')

    plt.xticks(x_pos, objects)    
    fig = plt.gcf()
    fig.set_size_inches(3,4)
    fig.tight_layout()   
    fig.savefig(directory + 'newcontentbar.png')
    plt.close(fig) 

def StandardBar(ReportDir, OVFilename, directory, daterange, type):
    print('Generating bar chart for new ' + type + '...')
    dfOv = pd.read_csv(ReportDir + "\\" + OVFilename) 

    dfOv = dfOv.sort_values(['Total number of new ' + type], ascending = False)
    dfOv = dfOv[dfOv['Total number of new ' + type] != 0]
    objects =  dfOv['PA'] 
    x_pos = np.arange(len(objects))

    if type == 'QandAs': dfOv['Total number of new ' + type].plot(kind='bar', facecolor='#bc8f8f')
    else: dfOv['Total number of new ' + type].plot(kind='bar', facecolor='#ECD44D')

    plt.subplots_adjust(bottom=0.5)
    plt.xticks(x_pos, objects)    
    
    fig = plt.gcf()
    fig.set_size_inches(6,4)
    fig.tight_layout()   
    if type == 'QandAs': fig.savefig(directory + 'newcontentbar-qas.png')
    if type == 'News items': fig.savefig(directory + 'newcontentbar-news.png')
    plt.close(fig) 
        
def Pie(listCT, directory, daterange):
    print('Generating pie chart for new content items...')
    
    labels = ['Practice Note', 'Overview', 'Checklist', 'Precedent']
    colors = ['#FF8AC4', '#91A8E8', '#ACFFBA', '#E8E2CE']
    explode = (0, 0, 0, 0) #change a zero to 0.1 to explode the piece you want
    matplotlib.rc('font', size=8) #set label font size
    matplotlib.rc('axes', titlesize=12) #set title font size
    
    def make_autopct(listCT):
        def my_autopct(pct):
            total = sum(listCT)
            val = int(round(pct*total/100.0))
            return '{v:d}'.format(p=pct,v=val) if pct > 0 else '' #'{p:.2f}%  ({v:d})'.format(p=pct,v=val)
        return my_autopct 

    plt.pie(listCT, explode=explode, labels=labels, colors=colors, autopct=make_autopct(listCT), pctdistance=0.8, labeldistance=10, startangle=90)#, autopct='%1.1f%%', startangle=270)
    plt.axis('equal')
    #plt.title(daterange)
    centre_circle = plt.Circle((0,0),0.70,fc='white')
    fig = plt.gcf()
    fig.gca().add_artist(centre_circle)
    fig.set_size_inches(3.5,3.5)
    fig.tight_layout()   
    plt.legend(loc=10)
    fig.savefig(directory + 'newcontentpie.png')
    plt.close(fig) 

def Export(menu, directory, ReportDir, outputfile, OVFilename, daterange, total, type, exportfilename):    
    style = '@import url("http://fonts.googleapis.com/css?family=Lato:400,700");th {    background-color: #69ab96;    color: #f5f5f0;}tr:nth-child(even) {background-color: #f2f2f2;}table, th, td, p, sup {   border: 0px solid black;   font-family: "Lato", sans-serif;   font-weight: 400;   text-align: left;   padding: 5px;   font-size: 14px;   border-spacing: 5px;   vertical-align: top;   text-align: left;}td {   word-wrap: break-word;   }tr:hover {background-color: #F8E0F7;}h1, h2, h3, h4, h5, h6 {   font-family: "Lato", sans-serif;   font-weight: 400;   }sup {    font-size: 10px;}img {    max-width: 100%;    display:block;    height: auto;}/* Style the tab */.tab {    float: left;    border: 1px solid #ccc;    background-color: #f1f1f1;    width: 16%;    height: 0px;}/* Style the buttons that are used to open the tab content */.tab button {    display: block;    background-color: inherit;    color: black;    padding: 5px 10px;    margin-right: 20px;    width: 100%;    border: none;    outline: none;    text-align: left;    cursor: pointer;    transition: 0.3s;    font-size: 13px;}/* Change background color of buttons on hover */.tab button:hover {    background-color: #E9FFEE;}/* Create an active/current tablink class */.tab button.active {    background-color: #ccc;}/* Style the tab content */.tabcontent {    position:absolute;    top: 15%;    left: 20%;    float: left;    padding: 0px 12px;    width: 100%;    border-left: none;    text-align: center;}.tabcontent {    animation: fadeEffect 1s; /* Fading effect takes 1 second */}.wrap{    width: 50%;    margin: 10px;}h2, img {    display: inline;    float: none;}/* Go from zero to full opacity */@keyframes fadeEffect {    from {opacity: 0;}    to {opacity: 1;}}#body {   padding:10px;   padding-bottom:60px;   /* Height of the footer */}.footer {    position: floating;    left: 0;    bottom: 10%;    width: 100%;}'
    print("Exporting to html...")
    pd.set_option('display.max_colwidth', -1) #stop the dataframe from truncating cell contents. This needs to be set if you want html links to work in cell contents
    df = pd.read_csv(ReportDir + outputfile, encoding='UTF-8') #read in the csv file
    dfOv = pd.read_csv(ReportDir + OVFilename)
    if type == 'qas': dfOv = dfOv[dfOv['Total number of new QandAs'] != 0] #filter by total to remove zero rows
    if type == 'news': 
        dfOv = dfOv[dfOv['Total number of new News items'] != 0] #filter by total to remove zero rows
        del df['Subtopic'], df['Content Type']
    if type == 'web': dfOv = dfOv[dfOv['Total number of new docs'] != 0] #filter by total to remove zero rows
    if type == 'email': dfOv = dfOv[dfOv['Total number of new docs'] != 0] #filter by total to remove zero rows
    dfOVHTML = dfOv.to_html(na_rep = " ", index=False)
    del df['DateFirstPublished']
    df['Link'] = '<a href="' + df['Link'] + '" target="_blank">View on PSL</a>' #mark up the link ready for html
    
    df.Link = df.Link.str.replace('—', ' ')
    df.Link = df.Link.str.replace('’',"'")
    dfHTML = df.to_html(na_rep = " ",index = False) #convert overview dataframe to html
    
    if type == 'web':
        html = r'<html><head><link rel="stylesheet" type="text/css" href="style.css"/><style></style>'
        html += '<title>New Content: ' + total + ' new docs created between ' + daterange + '</title></head>'
        html += '<h1 id="home">LexisPSL ContentHub: New Content</h1>' + menu + '<hr /><h2>' + total + ' new docs created between ' + daterange + '</h2>'
        html += '<p>This report includes the following content types only: PracticeNote, Overview, Checklist, and Precedents</p><hr />'
        html += '<div><img style="vertical-align: top" src="newcontentpie.png" /><img src="newcontentbar.png" /></div>'
        html += '<div style="overflow-x:auto;">' + dfHTML + '</div>' + '<hr />'
        html += '<p>LexisPSL New Content Report<br />Developed by Daniel Hutchings</p>'
        html = html.replace('&lt;', '<').replace('&gt;', '>').replace('\\', '/').replace('₂', '').replace('’',"'")
        
    if type == 'qas':
        total = str(len(df))
        html = r'<html><head><link rel="stylesheet" type="text/css" href="style.css"/><style></style></head>'
        html += '<title>New Q and As: ' + total + ' new Q and As created between ' + daterange + '</title>'
        html += '<h1 id="home">LexisPSL ContentHub: New Q and As</h1>' + menu + '<hr /><h2>' + total + ' new Q and As created between ' + daterange + '</h2>'
        html += '<p>This report includes only Q and As</p><hr />'
        html += '<div><img style="vertical-align: top" src="newcontentbar-qas.png" /></div>'
        html += '<div style="overflow-x:auto;">' + dfHTML + '</div>' + '<hr />'
        html += '<p>LexisPSL New Content Report QAs<br />Developed by Daniel Hutchings</p>'
        html = html.replace('&lt;', '<').replace('&gt;', '>').replace('\\', '/').replace('₂', '').replace('’',"'")
        
    if type == 'news':
        total = str(len(df))
        html = r'<html><head><link rel="stylesheet" type="text/css" href="style.css"/><style></style></head>'
        html += '<title>New News items: ' + total + ' new News items created between ' + daterange + '</title>'
        html += '<h1 id="home">LexisPSL ContentHub: New News items</h1>' + menu + '<hr /><h2>' + total + ' new News items created between ' + daterange + '</h2>'
        html += '<p>This report includes only News items</p><hr />'
        html += '<div><img style="vertical-align: top" src="newcontentbar-news.png" /></div>'
        html += '<div style="overflow-x:auto;">' + dfHTML + '</div>' + '<hr />'
        html += '<p>LexisPSL New Content Report News<br />Developed by Daniel Hutchings</p>'
        html = html.replace('&lt;', '<').replace('&gt;', '>').replace('\\', '/').replace('₂', '').replace('’',"'")
        
    if type == 'email':
        html = r'<html><head><link rel="stylesheet" type="text/css" href="style.css"/><style>' + style + '</style>'
        html += '<title>New Content: ' + total + ' new docs created between ' + daterange + '</title></head><body>'
        html += '<h1 id="home">LexisPSL ContentHub: New Content</h1><hr /><h2>' + total + ' new docs created between ' + daterange + '</h2>'
        html += '<p>This report includes the following content types only: PracticeNote, Overview, Checklist, and Precedents</p><hr />'
        html += '<div><img src="cid:image1" /><img src="cid:image2" /></div>'
        html += '<div style="overflow-x:auto;">' + dfHTML + '</div>' + '<hr />'
        html += '<p>LexisPSL New Content Report<br />Developed by Daniel Hutchings</p>'
        html += '</body></html>'
        html = html.replace('&lt;', '<').replace('&gt;', '>').replace('\\', '/').replace('₂', '').replace('’',"'")
        
    with open(directory + exportfilename,'w', encoding="utf-8") as f:
        f.write(html)
        f.close()
        pass
    
    print ("HTML generated at: " + directory + exportfilename)



def formatEmail(receiver_email, subject, filename):
    msg = MIMEMultipart("related")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = receiver_email

    # Encapsulate the plain and HTML versions of the message body in an
    # 'alternative' part, so message agents can decide which they want to display.
    msgAlternative = MIMEMultipart('alternative')
    msg.attach(msgAlternative)

    msgText = MIMEText('This is the alternative plain text message.')
    msgAlternative.attach(msgText)

    #send html from a file
    f = open(filename)
    msgText = MIMEText(f.read(),'html')
    # We reference the image in the IMG SRC attribute by the ID we give it below
    #msgText = MIMEText('<b>Some <i>HTML</i> text</b> and an image.<br><img src="cid:image1"><br><img src="cid:image2"><br>Nifty!', 'html')
    msgAlternative.attach(msgText)

    # This example assumes the image is in the current directory
    fp = open('C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\newcontentpie.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()

    # Define the image's ID as referenced above
    msgImage.add_header('Content-ID', '<image1>')
    msg.attach(msgImage)

    # This example assumes the image is in the current directory
    fp = open('C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\newcontentbar.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()

    # Define the image's ID as referenced above
    msgImage.add_header('Content-ID', '<image2>')
    msg.attach(msgImage)

    return msg

def sendEmail(msg, receiver_email):
    s = smtplib.SMTP("LNGWOKEXCP002.legal.regn.net")
    s.sendmail(sender_email, receiver_email, msg.as_string())

#main script
print("\nBuilding a list of the relevent AICER reports...")
menu = '<p><b>Currency</b>: <a href="Currency.html">General</a> | <a href="Top200Currency.html">Top 200</a> | <a href="ZeroCurrency.html">Zero Views</a> | <a href="ECC_Currency.html">Externally Commissioned</a> | <a href="BrexitCurrency.html">Brexit Related</a> | <a href="AGCurrency.html">AG</a> | <a href="CMSCurrency.html">CMS</a> | <a href="EVCurrency.html">EV</a> | <a href="IMCurrency.html">IM</a> | <a href="PMCurrency.html">PM</a>'
menu += '<br\><b>New Content</b>: <a href="Newcontentreport.html">General</a> | <a href="newcontentreportQandAs.html">QAs</a> | <a href="newcontentreportNews.html">News</a>'
menu += '<br\><b>Under Review</b>: <a href="underreviewsummary.html">General</a>'
menu += '<br\><b>Links</b>: <a href="file://///atlas/Knowhow/LinkHub/linkhub.html">LinkHub</a></p>'

directory2 = '\\\\atlas\\Knowhow\\ContentHub\\'
#directory2 = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\'
#lookupdpsi = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\'
lookupdpsi = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\lookupdpsi\\'
ReportDir = "\\\\atlas\\knowhow\\AICER\\reports\\"
#ReportDir = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\reports\\'
directory = "\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER_PM\\"
#directory = "\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER\\"
#directory = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\'

dirjoined = os.path.join(directory, '*AllContentItemsExport_*.csv') # this joins the directory variable with the filenames within it, but limits it to filenames ending in 'AllContentItemsExport_xxxx.csv', note this is hardcoded as 4 digits only. This is not regex but unix shell wildcards, as far as I know there's no way to specifiy multiple unknown amounts of numbers, hence the hardcoding of 4 digits. When the aicer report goes into 5 digits this will need to be modified, should be a few years until then though
files = sorted(glob.iglob(dirjoined), key=os.path.getctime, reverse=True) #search directory and add all files to dict
date =  str(time.strftime("%d/%m/%Y"))
weekago = (datetime.datetime.now().date() - datetime.timedelta(7)) #the 'date' part of this means it will only provide the date, not the hours, min, sec etc
OVFilename = 'new-content-overview.csv'
OVFilenameQAs = 'new-content-overview-QAs.csv'
OVFilenameNews = 'new-content-overview-News.csv'
emaildirectory = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\'
daterange = weekago.strftime('%b %d') + ' - ' + time.strftime('%b %d, %Y')
total = '0'

filename = files[0]

outputfile = re.search('.*\\\\AICER_PM\\\\([^\.]*)\.csv',filename).group(1) + "_UKPSL_newcontent.csv"
outputfileqandas = re.search('.*\\\\AICER_PM\\\\([^\.]*)\.csv',filename).group(1) + "_UKPSL_newcontent_QAs.csv"
outputfilenews = re.search('.*\\\\AICER_PM\\\\([^\.]*)\.csv',filename).group(1) + "_UKPSL_newcontent_News.csv"

filename = re.search('.*\\\\AICER_PM\\\\([^\.]*\.csv)',filename).group(1)   
print('\nLoaded: ' + files[0])


print('This is the most recent AICER report: ' + filename)
if os.path.isfile(ReportDir + outputfile) == False: #does the output file of this name already exist? If so skip
    Filter(filename, files, directory, ReportDir, outputfile, outputfileqandas, outputfilenews, weekago) #filter and categorise an aicer report
else: print('Filtered report already exists...skipping...')
    
MiniSummary(ReportDir, outputfile, lookupdpsi)
MiniSummary(ReportDir, outputfileqandas, lookupdpsi)
MiniSummary(ReportDir, outputfilenews, lookupdpsi)

OverviewLog(ReportDir, outputfile, OVFilename)
QandAsOverviewLog(ReportDir, outputfileqandas, OVFilenameQAs)
NewsOverviewLog(ReportDir, outputfilenews, OVFilenameNews)

StackedBar(ReportDir, OVFilename, emaildirectory, daterange)
StackedBar(ReportDir, OVFilename, directory2, daterange)

StandardBar(ReportDir, OVFilenameQAs, directory2, daterange, 'QandAs')
StandardBar(ReportDir, OVFilenameNews, directory2, daterange, 'News items')

df = pd.read_csv(ReportDir + outputfile, encoding='UTF-8') #Load csv file into dataframe
try: total = str(df.shape[0])
except: total = 0
try: pntotal = len(df[df['Content Type'] == 'PracticeNote'])
except: pntotal = 0
try: prtotal = len(df[df['Content Type'] == 'Precedent'])
except: prtotal = 0
try: cltotal = len(df[df['Content Type'] == 'Checklist'])
except: cltotal = 0
try: ovtotal = len(df[df['Content Type'] == 'AtAGlance'])
except: ovtotal = 0
listCT = [pntotal, ovtotal, cltotal, prtotal]

        
Pie(listCT, emaildirectory, daterange)
Pie(listCT, directory2, daterange)
Export(menu, directory2, ReportDir, outputfile, OVFilename, daterange, total, 'web', 'newcontentreport.html')  
Export(menu, emaildirectory, ReportDir, outputfile, OVFilename, daterange, total, 'email', 'newcontentreport_email.html') 
Export(menu, directory2, ReportDir, outputfileqandas, OVFilenameQAs, daterange, total, 'qas', 'newcontentreportQandAs.html')  
Export(menu, directory2, ReportDir, outputfilenews, OVFilenameNews, daterange, total, 'news', 'newcontentreportNews.html')    



#Email section
sender_email = 'LNGUKPSLDigitalEditors@ReedElsevier.com'
receiver_email_list = ['daniel.hutchings.1@lexisnexis.co.uk']
#receiver_email_list = ['daniel.hutchings.1@lexisnexis.co.uk', 'stephen.leslie@lexisnexis.co.uk', 'danielmhutchings@gmail.com', 'emma.millington@lexisnexis.co.uk', 'lisa.moore@lexisnexis.co.uk', 'claire.hayes@lexisnexis.co.uk', 'Ruth.Newman@lexisnexis.co.uk']

emaildirectory = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\'
#directory = '\\\\atlas\\Knowhow\\ContentHub\\'
filename = emaildirectory + 'newcontentreport_email.html'

tree = etree.parse(filename)
root = tree.getroot()
title = root.find('.//title')
subject = title.text


#create and send email
for receiver_email in receiver_email_list:
    msg = formatEmail(receiver_email, subject, filename)
    sendEmail(msg, receiver_email)
print('Email sent...')


print('Finished')
