#!/usr/bin/env python3

import xlwt
from xlwt import Workbook
import spacy
import operator
import numpy as np
import pandas as pd
from pandas import DataFrame
import json

# INITIALIZING SPACY AND ITS 'en' MODEL
nlp = spacy.load("en_core_web_sm")

# OPENING JSON SENTIMENT DICTIONARY
with open('afinn-165.json') as f:
  items_afinn = json.load(f)

#CLEANING DATA
f_stop_words=open("StopWords_GenericLong.txt", "r")
stop_words=[str(i[0:-1]) for i in f_stop_words]
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
df_sentiments = pd.DataFrame(sentiments_data, columns=['keyword', 'sentiment'])


create dataframe for materiality and keep count


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
    def tokenify_glossary(self,report):
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
        for sentence in sentences:
            sentence.split("but","yet","so")
            for section in sentence:
                doc=nlp(section)
                root = [token for token in doc if token.head == token][0]
                root=[root]
                tree=[]
                stakeholders = 0
                tree=create_tree(root,doc,tree)
                materiality = select_standards(doc)
                sentiment = calculate_sentiments(tree)




    #LIST-TREE
    def create_tree(self,temp_head,doc,TREE):
        doc = self.break_sentence(doc)
        doc = self.merge_compounds(doc)
        doc = self.combine_commas(doc)
        for sub_head in temp_head:
            self.track_subjects(sub_head)
            if not self.ignore_fillers(sub_head):
                BRANCH =[]
                BRANCH.append(sub_head.i)
                BRANCH.append(sub_head.dep_)
                BRANCH.append(sub_head.text)
                BRANCH.append(sub_head.head.i)
                TREE.append(BRANCH)
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
        for token in doc:
            if token.dep_=="compound" or (token.dep_=="conj" and (token.text!="and" or token.text!="or" or token.text!="nor")):
                index_c=token.i
                index_h=token.head.i
                with doc.retokenize() as retokenizer:
                    if index_h>index_c:
                        retokenizer.merge(doc[index_c:index_h+1], attrs={"dep":token.head.dep_})
                    elif index_c>index_h:
                        retokenizer.merge(doc[index_h:index_c+1], attrs={"dep":token.head.dep_})
                self.merge_compounds(doc)
            else:
                len=len+1
                if len >= doc[-1].i:
                    return doc

    #Combine descriptive
    def combine_commas(self,doc):
        len=0
        index_c=0
        index_h=0
        for token in doc:
            if (token.pos_=="ADV" or token.pos_=="ADJ" or token.pos_=="NOUN" or token.pos_=="VERB" or token.pos_=="PNON")and doc[token.i+1].text=="," and doc[token.i+2].pos_==token.pos_:
                index_c=token.i
                index_h=token.i+2
                with doc.retokenize() as retokenizer:
                    retokenizer.merge(doc[index_c:index_h+1], attrs={"dep":token.head.dep_})
                self.combine_commas(doc)
            else:
                len=len+1
                if len >= doc[-1].i:
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
            found=df_standards.loc[df_standards['sub-standard'].head(1) == token].values.tolist()
            if found:
                materiality.append([found[0][0], found[0][1]])
            elif found=df_standards.loc[df_standards['text'].head(1) == token].values.tolist():
                materiality.append(found[0][0], found[0][1]])
        return materiality


    #CALCULATE SENTIMENTS
    def calculate_sentiments(self,TREE):
        count_descriptive_words=0
        sentiment=0
        TREE=TREE.inverse()
        for items in TREE:
            item_sentiment=0
            items.split("and",",")
                for item in items:
                    found=df_sentiments.loc[df_sentiments['keyword'].head(1) == item].values.tolist()
                    if found:
                        item_sentiment=item_sentiment+1
                if not item_sentiment:
                    item_sentiment=1
            sentiment=item_sentiment=len(items)*sentiment
        return sentiments







    #CREATE GLORRARY
    # def sort_glossary(self,POS):
        # sorted_POS=[]
        #
        # unsorted_nouns = np.array(POS[0])
        # sorted_nouns=unsorted_nouns[unsorted_nouns[:, 1].argsort()]
        # sorted_nouns=self.reduce_glossary(sorted_nouns)
        # df_nouns = pd.DataFrame(sorted_nouns)
        # # df_nouns.columns=['frequency','text','lemma','pos','eng-tag','dependency','afinn sentiment','mcdonals sentiment','token id']
        # df_nouns.columns=['frequency','text','lemma','pos','afinn sentiment']
        #
        # unsorted_verbs = np.array(POS[1])
        # sorted_verbs=unsorted_verbs[unsorted_verbs[:, 1].argsort()]
        # sorted_verbs=self.reduce_glossary(sorted_verbs)
        # df_verbs = pd.DataFrame(sorted_verbs)
        # df_verbs.columns=['frequency','text','lemma','pos','afinn sentiment']
        #
        # unsorted_adverbs = np.array(POS[2])
        # sorted_adverbs=unsorted_adverbs[unsorted_adverbs[:, 1].argsort()]
        # sorted_adverbs=self.reduce_glossary(sorted_adverbs)
        # df_adverbs = pd.DataFrame(sorted_adverbs)
        # df_adverbs.columns=['frequency','text','lemma','pos','afinn sentiment']
        #
        # unsorted_adjectives = np.array(POS[3])
        # sorted_adjectives=unsorted_adjectives[unsorted_adjectives[:, 1].argsort()]
        # # sorted_adjectives=self.group_synoynms(self.reduce_glossary(sorted_adjectives))
        # sorted_adjectives=self.reduce_glossary(sorted_adjectives)
        # df_adjective = pd.DataFrame(sorted_adjectives)
        # df_adjective.columns=['frequency','text','lemma','pos','afinn sentiment']
        #
        # sorted_POS.append(sorted_nouns)
        # sorted_POS.append(sorted_verbs)
        # sorted_POS.append(sorted_adverbs)
        # sorted_POS.append(sorted_adjectives)

        #glossary to excel
        # with pd.ExcelWriter("companies_glossary/"+self.company+".xlsx") as writer:
        #     df_nouns.to_excel(writer, sheet_name='Nouns')
        #     df_verbs.to_excel(writer, sheet_name='Verbs')
        #     df_adverbs.to_excel(writer, sheet_name='Adverbs')
        #     df_adjective.to_excel(writer, sheet_name='Adjectives')
        # writer.save()
        #
        # return sorted_POS
    #print(sort_glossary(divide_glossary(tokenify_glossary(read_file()))))

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
