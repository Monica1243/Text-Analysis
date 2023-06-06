import pandas as pd
import re
import openpyxl
import os
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords

input_file = openpyxl.load_workbook("Input.xlsx")

column_used = input_file.active

file1 = "StopWords\StopWords_Auditor.txt"
file2 = "StopWords\StopWords_Currencies.txt"
file3 = "StopWords\StopWords_DatesandNumbers.txt"
file4 = "StopWords\StopWords_Generic.txt"
file5 = "StopWords\StopWords_GenericLong.txt"
file6 = "StopWords\StopWords_Geographic.txt"
file7 = "StopWords\StopWords_Names.txt"
file8 = "MasterDictionary/negative_words.txt"
file9 = "MasterDictionary/positive_words.txt"

with open(file1,"r") as f1:
        auditor = [word.strip() for word in f1.readlines()]
with open(file2,"r") as f2:
    currencies = [word.strip() for word in f2.readlines()]
with open(file3,"r") as f3:
    datesAndNumbers = [word.strip() for word in f3.readlines()]
with open(file4,"r") as f4:
    generic = [word.strip() for word in f4.readlines()]
with open(file5,"r") as f5:
    genericLong = [word.strip() for word in f5.readlines()]
with open(file6,"r") as f6:
    geographic = [word.strip() for word in f6.readlines()]
with open(file7,"r") as f7:
    names = [word.strip() for word in f7.readlines()]
with open(file8,"r") as f8:
    negative = [word.strip() for word in f8.readlines()]
with open(file9,"r") as f9:
    positive = [word.strip()   for word in f9.readlines()]
        
all_words = auditor + currencies + datesAndNumbers + generic + genericLong + geographic + names

def stopWords(title):
    title_split = title.split()
    analyse_words = [word for word in title_split if word not in all_words]

    analysis_words = " ".join(analyse_words)
    return analysis_words

def segment_analysis(args):
    words = stopWords(args)
    length = len(words)
    positive_words = 0
    negative_words  = 0
    
    for word in word_tokenize(words):
        if word in negative:
            negative_words += 1
        if word in positive:
            positive_words += 1

    polarity = (positive_words - negative_words)/ ((positive_words + negative_words) + 0.000001)
    subjectivity = (positive_words + negative_words)/(length + 0.000001)

    return negative_words , positive_words, polarity, subjectivity
    
def word_count(args):
        stops = set(stopwords.words('english'))
        words = word_tokenize(args)
        wordsFiltered = []
        for w in words:
            if w not in stops:
                wordsFiltered.append(w)
        print(' '.join(wordsFiltered))
        print(len(wordsFiltered))
                
regex = r"(?<=\.com\/)[^\/]*"

for row in column_used.iter_rows(values_only=True):
    url = row[1]
    title = re.findall(regex,url)
    if title:
        for value in title:
            args = value.replace("-"," ")
            negative_words , positive_words, polarity, subjectivity = segment_analysis(args)
            word_count(args)
            
##            print(negative_words , positive_words, polarity, subjectivity)
