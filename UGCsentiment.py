# import tweepy
#
# import xlwt
# from xlwt import Workbook
#
# import json
# with open('afinn.json') as f:
#   data = json.load(f)
#
# wb = Workbook(encoding='ascii')
# sheet1 = wb.add_sheet('Tweets')
# sheet2 = wb.add_sheet('Sentiment')
# # Authenticate to Twitter
# auth = tweepy.OAuthHandler("oG**********56", "hoq**********************llGK")
# auth.set_access_token("847**************LG9E", "ek*********************qp14")
# api = tweepy.API(auth) # Create API object
#
# f_glossary=open("glossary.txt", "r")
# fin = []
# soc = []
# env = []
# item_g = f_glossary.read().split("|")
# fin=item_g[0].split(",")
# soc=item_g[1].split(",")
# env=item_g[2].split(",")
# standards=[env,soc,fin]
# standards_title=['env','soc','fin']
# for j in range(3):
#     sheet1.write(0, j, standards_title[j])
#     count = 1
#     for k in range(10):
#         search_terms=[standards[j][k], "HUL"]
#         public_tweets = api.search(q=search_terms, lang="en", count=100)
#     # for each through all tweets pulled # q= “string that we are looking for”
#         for tweet in public_tweets:
#             sheet1.write(count, j, tweet.text)  #print(standards[j][k],count)
#             sheet1.write(count, j+3, str(tweet.created_at))
#             print (tweet.created_at)   #printing the time
#             sentence=tweet.text.split(' ')
#             sentiment_count =0
#             for l in sentence:
#                 if l in data:
#                     sentiment_count = sentiment_count + data[l]
#             sheet2.write(count, j, sentiment_count/len(tweet.text))
#             count=count+1
#     print(count)
# wb.save('Tweets.xls')
#
#
#
#





















import tweepy

import xlwt
from xlwt import Workbook

import json

with open('afinn.json') as f:
  data = json.load(f)

wb = Workbook(encoding='ascii')
sheet1 = wb.add_sheet('Tweets')
sheet2 = wb.add_sheet('Sentiment')


# Authenticate to Twitter
auth = tweepy.OAuthHandler("oGMSF0ZE60EPmq63CnJ5OVy56", "hoqrp3VgWWvzE3OjAT6iCD8upY4vc3UwcswH3Fk5Q4x46xllGK")
auth.set_access_token("847417106-WmqvtW8uv0HSHy0MXMpOjLEesVHoq450uXODLG9E", "ek7td0b383tqC75ZZ3bcuOqzl9ooxOrkpUAdjjHs5qp14")

# Create API object
api = tweepy.API(auth)

f_glossary=open("glossary.txt", "r")
fin = []
soc = []
env = []
item_g = f_glossary.read().split("|")
fin=item_g[0].split(",")
soc=item_g[1].split(",")
env=item_g[2].split(",")
standards=[env,soc,fin]
standards_title=['env','soc','fin']
#print(standards)
# Create a tweet
#api.update_status("Hello Tweepy")

# Using the API object to get tweets from your timeline, and storing it in a variable called public_tweets
#public_tweets = api.home_timeline()

# Using the API object to get tweets from a user’s timeline, and storing it in a variable called public_tweets
#public_tweets = api.user_timeline(id,count)     # max is 20 coz twitter put a limit to it
for j in range(3):
    sheet1.write(0, j, standards_title[j])
    count = 1
    for k in range(10):
        search_terms=[standards[j][k], "HUL"]
#search_terms=["plastic HUL", "air HUL", "water HUL", "energy HUL", "materials HUL", "waste HUL", "habitat HUL", "wildlife HUL", "sustainable HUL", "emissions HUL", "plastic HUL"]
    #for i in range(len(search_terms)):
        public_tweets = api.search(q=search_terms, lang="en", count=100)      # q= “string that we are looking for”
        # foreach through all tweets pulled
        for tweet in public_tweets:
            #print(standards[j][k],count)
            sheet1.write(count, j, tweet.text)
            sheet1.write(count, j+3, str(tweet.created_at))
            #print (tweet.text)   #printing the text stored inside the tweet object
            print (tweet.created_at)   #printing the time
            sentence=tweet.text.split(' ')
            sentiment_count =0
            for l in sentence:
                if l in data:
                    sentiment_count = sentiment_count + data[l]
            sheet2.write(count, j, sentiment_count/len(tweet.text))
            count=count+1
    print(count)
    #print (tweet.user.screen_name)   #printing the username
    #print (tweet.user.location)   #printing the location

# API reference index – Twitter developer tools
#https://developer.twitter.com/en/docs/api-reference-index

# Rate Limiting - Twitter
#https://developer.twitter.com/en/docs/basics/rate-limiting

wb.save('Tweets.xls') 
