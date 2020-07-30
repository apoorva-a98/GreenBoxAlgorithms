#!/usr/bin/env python3

import tweepy
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
  data = json.load(f)

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

glossary=standards_data.values.tolist()
keywords=[]
for item in glossary:
    if item[2] not in keywords:
        keywords.append(item[2])
    if item[1] not in keywords:
        keywords.append(item[1])

# Authenticate to Twitter
auth = tweepy.OAuthHandler("oGMSF0ZE60EPmq63CnJ5OVy56", "hoqrp3VgWWvzE3OjAT6iCD8upY4vc3UwcswH3Fk5Q4x46xllGK")
auth.set_access_token("847417106-WmqvtW8uv0HSHy0MXMpOjLEesVHoq450uXODLG9E", "ek7td0b383tqC75ZZ3bcuOqzl9ooxOrkpUAdjjHs5qp14")

# Create API object
api = tweepy.API(auth)
# api.update_status(status='Test')


class reports:
    def __init__(self, name):
        self.company = name

    # TOKANIZING SINGLE TWEET
    def tokenify_tweet(self,tweet):
        buff = ''
        sentences=[]
        for letter in tweet:
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

    #READING TWEETS
    def read_tweets(self):
        master_list=pd.DataFrame(columns=['standards','sub-standard','sentence', 'sentiment','time','retweets'])
        stakeholders = 0
        materiality_count = []
        for keyword in keywords:
            search_terms=[keyword, self.company]
            # public_tweets = api.search(q=search_terms, lang="en", count=100)      # q= “string that we are looking for”
            # public_tweets=[]
            for tweet in tweepy.Cursor(api.search, q=search_terms, lang="en").items(5000000):
                if (not tweet.retweeted) and ('RT @' not in tweet.text):
            #         try:
            #             item=[]
            #             item.append(tweet.text.encode('utf-8'))
            #             item.append(tweet.created_at)
            #             item.append(tweet.retweet_count)
            #             public_tweets.append(item)
            #         except:
            #             continue
            # print(public_tweets)

                # for each tweet within all tweets pulled
                # for tweet in public_tweets:
                    tweet_time="-"
                    tweet_retweet_count="-"
                    try:
                        tweet_time=tweet.created_at
                        tweet_retweet_count=tweet.retweet_count
                    except:
                        pass
                    tweet=str(tweet.text.encode('utf-8'))
                    sentences = self.tokenify_tweet(tweet)
                    for sentence in sentences:
                        doc=nlp(sentence)
                        try:
                            root = [token for token in doc if token.head == token][0]
                        except:
                            continue
                        root=[root]
                        tree=[]
                        materiality = self.select_standards(doc)
                        if materiality:
                            tree=self.create_tree(root,doc,tree)
                            sentiment = self.calculate_sentiments(tree)
                            for sub_standard in materiality:
                                # print(sub_standard,materiality,len(materiality))
                                # print(root,doc,sub_standard[0],sub_standard[1],sentence,sentiment)
                                data_point=[]
                                data_point.append(sub_standard[0])
                                data_point.append(sub_standard[1])
                                data_point.append(sentence)
                                data_point.append(sentiment)
                                data_point.append(tweet_time)
                                data_point.append(tweet_retweet_count)
                                to_append = data_point
                                df_length = len(master_list)
                                master_list.loc[df_length] = to_append

                                if sub_standard[1] not in materiality_count:
                                    materiality_count.append(sub_standard[1])

        print(stakeholders)
        print(materiality_count)
        return master_list


    #LIST-TREE
    def create_tree(self,temp_head,doc,TREE):
        for sub_head in temp_head:
            self.track_subjects(sub_head)
            if not self.ignore_fillers(sub_head):
                TREE.append(sub_head)
            sub_tree=[child for child in sub_head.children]
            if not sub_tree:
                continue
            else:
                self.create_tree(sub_tree,doc,TREE,)
        return TREE

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
            found=standards_data.loc[standards_data['sub-standard'] == token].head(1).values.tolist()
            if found:
                sub_standard=[]
                sub_standard.append(found[0][0])
                sub_standard.append(found[0][1])
                materiality.append(sub_standard)
                break
            else:
                found=standards_data.loc[standards_data['text'] == token.lemma_].head(1).values.tolist()
                if found:
                    sub_standard=[]
                    sub_standard.append(found[0][0])
                    sub_standard.append(found[0][1])
                    materiality.append(sub_standard)
                else:
                    continue
        if not materiality:
            return 0
        return materiality

    #CALCULATE SENTIMENTS
    def calculate_sentiments(self,TREE):
        #add afinn
        count_descriptive_words=0
        sentiment=0
        for item in reversed(TREE):
            item_sentiment=0
            try:
                found=sentiments_data.loc[sentiments_data["keyword"] == item.lemma_].head(1).values.tolist()
                if found:
                    item_sentiment=item_sentiment+found[0][1]
                    count_descriptive_words=count_descriptive_words+1
                elif item in data:
                    item_sentiment=item_sentiment+data[item]
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

    # CREATE DATABASE
    def create_database(self,df_master_list):

        # database to excel
        with pd.ExcelWriter(self.company+"tweets.xlsx") as writer:
            df_master_list.to_excel(writer)
        writer.save()

HUL = reports("HUL")
print(HUL.create_database(HUL.read_tweets()))
