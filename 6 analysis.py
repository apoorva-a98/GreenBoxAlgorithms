#!/usr/bin/env python3

import xlwt
from xlwt import Workbook
import spacy
import operator
import numpy as np
import pandas as pd
from pandas import DataFrame
import json
import re

# INITIALIZING SPACY AND ITS 'en' MODEL
nlp = spacy.load("en_core_web_sm")

# OPENING JSON SENTIMENT DICTIONARY
with open('afinn-165.json') as f:
  items_afinn = json.load(f)

#CLEANING DATA
# f_stop_words=open("StopWords_GenericLong.txt", "r")
# stop_words=[str(i[0:-1]) for i in f_stop_words]
avoid=['@','#','$','%','^','&','*','(',')','_','=','+','[',']','|','\n','\t','<','>','/']

# READING GLOSSARY EXCEL
def read_standards():
    sheet= pd.read_excel("Standards.xlsx")
    return sheet

# READING SENTIMENTS EXCEL
def read_sentiments():
    sheet= pd.read_excel("Sentiments.xlsx")
    return sheet

standards_data=read_standards()
sentiments_data=read_sentiments()
df_standards = pd.DataFrame(standards_data)
df_sentiments = pd.DataFrame(sentiments_data)
# df_sentiments = pd.DataFrame(sentiments_data, columns=['keyword', 'sentiment'])
# print(df_standards)
# print(df_sentiments)


found=df_standards.loc[df_standards['sub-standard'] == "water"].head(1).values.tolist()
if found:
    print([found[0][0], found[0][1]])

found=df_sentiments.loc[df_sentiments['keyword'] == "attempt"].head(1).values.tolist()
if found:
    print([found[0][0], found[0][1]])


class reports:
    def __init__(self, name, path):
        self.company = name
        self.filepath = path
        subject=0


    # READING COMPANY REPORTS
    def read_file(self):
        report= open(self.filepath, "r", encoding = "ISO-8859-1")
        report_text= report.read()
        return report_text
    #print(read_file())


    # TOKANIZING REPORT
    def tokenify_sentences(self,report):
        buff = ''
        sentences=[]
        for letter in report:
            letter=letter.lower()
            if letter in avoid:
                if buff != '':
                    sentences.append(buff)
                buff = ''
            elif (buff is not None):
                buff += letter
        if buff is not None:
            sentences.append(buff)
            buff=''
        return sentences
    #print(tokenify_glossary(read_file()))


    #READ SENTENCES
    def read_sentences(self,sentences):
        master_list=pd.DataFrame(columns=['standards','sub-standard','sentence', 'sentiment'])
        for sentence in sentences:
            # sentence=re.split('but yet so',sentence)
            # sentence.split("but","yet","so")
            # for section in sentence:
            doc=nlp(sentence)
            try:
                root = [token for token in doc if token.head == token][0]
            except:
                continue
            root=[root]
            tree=[]
            stakeholders = 0
            materiality = self.select_standards(doc)
            if materiality:
                tree=self.create_tree(root,doc,tree)
                sentiment = self.calculate_sentiments(tree)
                for sub_standard in materiality:
                    # print(root,doc,sub_standard[0],sub_standard[1],sentence,sentiment)
                    data_point=[]
                    data_point.append(sub_standard[0])
                    data_point.append(sub_standard[1])
                    data_point.append(sentence)
                    data_point.append(sentiment)
                    to_append = data_point
                    df_length = len(master_list)
                    master_list.loc[df_length] = to_append
                    # print(data_point)
                    # df_data_point=pd.DataFrame(data_point)
                    # master_list=master_list.append(df_data_point)
        return master_list


    #LIST-TREE
    def create_tree(self,temp_head,doc,TREE):
        # doc = self.break_sentence(doc)
        # doc = self.merge_compounds(doc)
        # doc = self.combine_commas(doc)
        for sub_head in temp_head:
            self.track_subjects(sub_head)
            if not self.ignore_fillers(sub_head):
                # BRANCH =[]
                # BRANCH.append(sub_head.i)
                # BRANCH.append(sub_head.dep_)
                # BRANCH.append(sub_head.text)
                # BRANCH.append(sub_head.head.i)
                # TREE.append(BRANCH)
                TREE.append(sub_head)
            sub_tree=[child for child in sub_head.children]
            if not sub_tree:
                continue
            else:
                self.create_tree(sub_tree,doc,TREE,)
        return TREE

    #Merge dep
    def merge_compounds(self,doc):
        len=0
        index_c=0
        index_h=0
        index_max=doc[-1].i
        buffer=doc
        for token in doc:
            if token.dep_=="compound" or (token.dep_=="conj" and (token.text!="and" or token.text!="or" or token.text!="nor")):
                index_c=token.i
                index_h=token.head.i
                with doc.retokenize() as retokenizer:
                    if index_h>index_c:
                        retokenizer.merge(doc[index_c:index_h+1], attrs={"dep":token.head.dep_})
                    elif index_c>index_h:
                        retokenizer.merge(doc[index_h:index_c+1], attrs={"dep":token.head.dep_})
                if doc!=buffer:
                    self.merge_compounds(doc)
                else:
                    continue
            else:
                len=len+1
                if len >= index_max:
                    return doc

    #Combine descriptive
    def combine_commas(self,doc):
        len=0
        index_c=0
        index_h=0
        if doc[-1].i>=3:
            index_max=doc[-1].i
            for token in doc:
                if (token.pos_=="ADV" or token.pos_=="ADJ" or token.pos_=="NOUN" or token.pos_=="VERB" or token.pos_=="PNON")and doc[token.i+1].text=="," and doc[token.i+2].pos_==token.pos_:
                    index_c=token.i
                    index_h=token.i+2
                    with doc.retokenize() as retokenizer:
                        if index_h<index_max-1:
                            retokenizer.merge(doc[index_c:index_h+1], attrs={"dep":token.head.dep_})
                    self.combine_commas(doc)
                else:
                    len=len+1
                    if len >= index_max:
                        return doc

    #Address sentence breaks
    def break_sentence(self,doc):
        for token in doc:
            if (token.dep_=="conj" and (token.text!="and" or token.text!="or" or token.text!="nor")) :
                token.text=","
        return doc

    #Track subjects and pronouns
    def track_subjects(self,sub_head):
        if sub_head.dep_=="nsubj" or sub_head.dep_=="csubj":
            subject=sub_head.text
        elif sub_head.dep_=="PRON" and (sub_head.text=="It" or sub_head.text=="it"):
            sub_head.text=subject
        elif sub_head.dep_=="PRON" and (sub_head.text!="It" or sub_head.text!="it"):
            stakeholders=stakeholders+1

    #Ignore dep
    def ignore_fillers(self,sub_head):
        if sub_head.dep_=="aux" or sub_head.pos_=="DET" or sub_head.pos_=="PUNCT" or sub_head.dep_=="preconj" or sub_head.dep_=="prep":
            return 1
        else:
            return 0


    # INDENTIFY MATERIALITY
    def select_standards(self,doc):
        materiality=[]
        for token in doc:
            found=df_standards.loc[df_standards['sub-standard'] == token].head(1).values.tolist()
            # try:
            if found:
                # materiality.append([found[0][0], found[0][1]])
                sub_standard=[]
                sub_standard.append(found[0][0])
                sub_standard.append(found[0][1])
                materiality.append(sub_standard)
                break
            else:
                found=df_standards.loc[df_standards['text'] == token.lemma_].head(1).values.tolist()
                if found:
                    # materiality.append([found[0][0], found[0][1]])
                    sub_standard=[]
                    sub_standard.append(found[0][0])
                    sub_standard.append(found[0][1])
                    materiality.append(sub_standard)
                else:
                    continue
            # except:
                # pass
        if not materiality:
            return 0
        return materiality

    #CALCULATE SENTIMENTS
    def calculate_sentiments(self,TREE):
        count_descriptive_words=0
        sentiment=0
        for item in reversed(TREE):
            item_sentiment=0
            try:
                found=df_sentiments.loc[df_sentiments["keyword"] == item.lemma_].head(1).values.tolist()
                if found:
                    item_sentiment=item_sentiment+found[0][1]
                    count_descriptive_words=count_descriptive_words+1
            except TypeError:
                pass

            if not item_sentiment:
                item_sentiment=1
            if str(item)=="not" or str(item)=="nor":
                print(item)
                sentiment=item_sentiment+(-1)*sentiment
            else:
                sentiment=item_sentiment+sentiment
        # print(sentiment,count_descriptive_words)
        if count_descriptive_words:
            sentiment=sentiment/count_descriptive_words
        return sentiment


    # #CALCULATE SENTIMENTS
    # def calculate_sentiments(self,TREE):
    #     count_descriptive_words=0
    #     sentiment=0
    #     # TREE=TREE.inverse()
    #     for items in reversed(TREE):
    #         item_sentiment=0
    #         try:
    #             items=re.split('and ,',items)
    #             for item in items:
    #                 found=df_sentiments.loc[df_sentiments['keyword'] == item].head(1).values.tolist()
    #                 if found:
    #                     item_sentiment=item_sentiment+1
    #                     count_descriptive_words=count_descriptive_words+1
    #         except TypeError:
    #             pass
    #         # items.split("and",",")
    #         if not item_sentiment:
    #             item_sentiment=1
    #         sentiment=item_sentiment+len(items)*sentiment
    #     return sentiment


    # CREATE DATABASE
    def create_database(self,df_master_list):

        # database to excel
        with pd.ExcelWriter(self.company+"database.xlsx") as writer:
            df_master_list.to_excel(writer)
        writer.save()


HUL = reports("HUL", "HUL 2018-2019_Annual Report.txt")
# print(HUL.read_sentences(HUL.tokenify_sentences(HUL.read_file())))
print(HUL.create_database(HUL.read_sentences(HUL.tokenify_sentences(HUL.read_file()))))










# HUL = reports("HUL", "HUL 2018-2019_Annual Report.txt")
# print(HUL.sort_glossary(HUL.divide_glossary(HUL.tokenify_glossary(HUL.read_file()))))
#
# Colgate = reports("Colgate", "Colgate 2018-2019_Annual Report.txt")
# print(Colgate.sort_glossary(Colgate.divide_glossary(Colgate.tokenify_glossary(Colgate.read_file()))))
#
# ITC = reports("ITC", "ITC 2018-2019 Annual Report.txt")
# print(ITC.sort_glossary(ITC.divide_glossary(ITC.tokenify_glossary(ITC.read_file()))))
#
# Dabur = reports("Dabur", "Dabur 2018-19_Annual Report.txt")
# print(Dabur.sort_glossary(Dabur.divide_glossary(Dabur.tokenify_glossary(Dabur.read_file()))))
#
# Godrej = reports("Godrej", "Godrej 2018-2019_Annual Report.txt")
# print(Godrej.sort_glossary(Godrej.divide_glossary(Godrej.tokenify_glossary(Godrej.read_file()))))
#
# Marico = reports("Marico", "Marico 2018-2019_Annual Report.txt")
# print(Marico.sort_glossary(Marico.divide_glossary(Marico.tokenify_glossary(Marico.read_file()))))
#
# Nestle = reports("Nestle", "Nestle 2017-2018_Annual Report.txt")
# print(Nestle.sort_glossary(Nestle.divide_glossary(Nestle.tokenify_glossary(Nestle.read_file()))))
#
# PnG = reports("PnG", "P&G 2018-2019_Annual Report.txt")
# print(PnG.sort_glossary(PnG.divide_glossary(PnG.tokenify_glossary(PnG.read_file()))))
