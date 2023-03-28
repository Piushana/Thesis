# -*- coding: utf-8 -*-
"""
Created on Tue Mar 28 13:14:35 2023

@author: piushana
"""

# Import the necessary libraries
from googleapiclient.discovery import build
from google.oauth2 import service_account
import json
import csv
from ibm_watson import NaturalLanguageUnderstandingV1
from ibm_cloud_sdk_core.authenticators import IAMAuthenticator
from ibm_watson.natural_language_understanding_v1 \
import Features, EntitiesOptions, KeywordsOptions, EmotionOptions
import pandas as pd  

## functions 
## extract answers from dictionary  
def findAnswer(ResponsesList, ID, key):
    return ResponsesList[ID]['answers'][key]['textAnswers']['answers'][0]['value']


# Apply color to cells based on their values
# tone.val<0.5(green), tone.val0.5-075(blue), tone.val>0.75(red)
def colorize(val):
    if val < 0.5:
        return 'background-color: green'
    elif val < 0.75:
        return 'background-color: blue'
    else:
        return 'background-color: red'


## Tone analysis with IBM NLU
def toneAnalyze(Answer):
    authenticator = IAMAuthenticator('QY3XwG1AXzsZVqZd3to2cMECzVlRgOQKKwSP1WhTDFwz')
    natural_language_understanding = NaturalLanguageUnderstandingV1(
       version='2022-04-07',
        authenticator=authenticator)
    natural_language_understanding.set_service_url('set_service_url')
    
    response = natural_language_understanding.analyze(
        text=Answer,
         language="en",
       features=Features(emotion=EmotionOptions())).get_result()
    
    return response

# Set the form ID of the goole form to retrive data
form_id = 'form_id'

# Get auth 2.0 credentials for authetication to forms
creds = service_account.Credentials.from_service_account_file('set_authentication_credentials')


# Google Forms API service
service = build('forms', 'v1', credentials=creds)

# Use the service to get the form responses
Responses = service.forms().responses().list(formId=form_id).execute()
ResponsesList = Responses['responses']

#call the function findAnswer()
#'6fbd522c'-Question_id
tones = []
for x in range (40):
    answer1 = findAnswer (ResponsesList, x, '6fbd522c')
    toneAnswer = toneAnalyze(answer1)
    toneAnswer = toneAnswer['emotion']['document']['emotion']
    toneAnswer['tone'] = max(toneAnswer.keys(), key=lambda k: toneAnswer[k])
    toneAnswer['Answer'] = answer1
    tones.append(toneAnswer)

    
field_names = ['Answer', 'sadness', 'joy','fear','disgust','anger','tone']
# open CSV file
filename = 'data.csv'
with open(filename, 'w', newline='') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=field_names)
    # write the header rows
    writer.writeheader()
    # write the data rows
    for row in tones:
        writer.writerow(row)

# for the analysis section to add colours according to above mentioned conditions (green- no tone/ blue-tone detected/ red-strong tone)
# Read the CSV file into a DataFrame
df = pd.read_csv('data.csv')
# Apply the colors to the mentioned columns
styled_df = df.style.applymap(colorize, subset=['sadness', 'joy', 'fear', 'disgust', 'anger'])

# Save the coloured DataFrame as an Excel workbook
styled_df.to_excel('Final.xlsx', engine='openpyxl', index=False)


