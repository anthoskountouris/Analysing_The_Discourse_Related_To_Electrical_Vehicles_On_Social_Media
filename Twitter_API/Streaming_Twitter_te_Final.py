# https://docs.tweepy.org/en/v3.9.0/streaming_how_to.html
# https://developer.twitter.com/en/docs/twitter-api/data-dictionary/object-model/tweet
# https://docs.tweepy.org/en/v3.9.0/extended_tweets.html
# https://developer.twitter.com/en/docs/twitter-api/v1/tweets/filter-realtime/guides/basic-stream-parameters
# https://developer.twitter.com/en/docs/twitter-api/v1/tweets/filter-realtime/api-reference/post-statuses-filter
# https://www.programiz.com/python-programming/methods/built-in/hasattr
# https://stackoverflow.com/questions/31968069/unicode-characters-in-twitter-python
# https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file
# https://stackoverflow.com/questions/22889122/how-to-add-a-location-filter-to-tweepy-module

# pip install --upgrade tweepy==3.8.0

import tweepy
#import twitter keys and tokens
import API_keys
import openpyxl


# Keys and Access Tokens
consumer_key =  API_keys.consumer_key
consumer_secret = API_keys.consumer_secret
access_token =  API_keys.access_token
access_token_secret =  API_keys.access_token_secret

# keywords lists
keywords = ['electric vehicles', '#electricvehicles']
keywords_EV = ['evs', 'ev','#ev','#evs','ev\'s']


class MyStreamListener(tweepy.StreamListener):
    def __init__(self, api):
        self.api = api
        self.me = api.me()

    def on_connect(self):
        print("Connected to Twitter API ")

    def on_status(self, status):

        #open excel the document 
        wb = openpyxl.load_workbook('Twitter_Dataset.xlsx')

        hasLocation = False
        hasWantedKeyword = False

        """This Status event handler for a StreamListener prints the full text of the Tweet, 
        or if it’s a Retweet, the full text of the Retweeted Tweet

        If status is a Retweet, it will not have an extended_tweet attribute, and status.text could be truncated."""

        if hasattr(status, "retweeted_status"):  # Check if Retweet
            try:
                tweet = status.retweeted_status.extended_tweet["full_text"]
            except AttributeError:
                tweet = status.retweeted_status.text
        else:
            try:
                tweet = status.extended_tweet["full_text"]
            except AttributeError:
                tweet = status.text

        # if tweet has the place object is True
        if status.place is not None:
            hasLocation = True
            print ("The status has location")

        tweet_lowercase = tweet.lower()
        words = tweet.lower().split()

        #checking whether the tweer has one of the keywords
        if  any(word in words for word in keywords_EV):
            hasWantedKeyword = True
            print("The tweet has the keyword_EV")
        elif any(word in tweet_lowercase for word in keywords):
            hasWantedKeyword = True
            print("The tweet has a keyword")

        #if both location and keyword founs then retrieveing the tweet    
        if hasLocation and hasWantedKeyword:
            excel_input_list = []

            tweet = str(tweet)
            print (">> " + tweet)

            location_of_tweet = str(status.place.full_name)
            print (location_of_tweet)

            date_created = str(status.created_at)
            print (date_created)

            excel_input_list.append(tweet)
            excel_input_list.append(location_of_tweet)
            excel_input_list.append(date_created)

            print(excel_input_list)

            try:
                #Appending the excel_input_list in the xlsx file
                ws = wb['Sheet1']
                ws.append(excel_input_list) 
                wb.save('Twitter_Dataset.xlsx')
            except:
                print("Appending to file error")


    def on_error(self, status):
        print("Error detected")

# Authenticate to Twitter
auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
auth.set_access_token(access_token, access_token_secret)

# Create API object
api = tweepy.API(auth, wait_on_rate_limit=True,
    wait_on_rate_limit_notify=True)

# Once we have an api and a status listener we can create our stream object
tweets_listener = MyStreamListener(api)
stream = tweepy.Stream(api.auth, tweets_listener)

#Starting the stream using the filter to stream the tweets that are posted in thr bounding box of the UK
try:
    stream.filter(locations=[-6.38,49.87,1.77,55.81], languages=["en"])
except KeyboardInterrupt:
    print('The stream has stopped')
finally:
    stream.disconnect()


