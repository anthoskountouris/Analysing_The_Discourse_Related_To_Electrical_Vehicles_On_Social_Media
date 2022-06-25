# https://docs.tweepy.org/en/v4.5.0/client.html#tweepy.Client.search_recent_tweets
# https://developer.twitter.com/en/docs/twitter-api/tweets/search/integrate/build-a-query#list
# https://developer.twitter.com/en/docs/twitter-api/tweets/search/integrate/build-a-query
# https://developer.twitter.com/en/docs/twitter-api/data-dictionary/object-model/place
# https://docs.tweepy.org/en/v4.5.0/client.html#place-fields-parameter
# https://dev.to/twitterdev/a-comprehensive-guide-for-using-the-twitter-api-v2-using-tweepy-in-python-15d9
# https://www.youtube.com/watch?v=rQEsIs9LERM


import tweepy
#import twitter keys and tokens
import config
import random
import json
import openpyxl

# Getting the Client with the Keys and Access Tokens
def getClient():
	client = tweepy.Client(bearer_token=config.bearer_token, 
						   consumer_key=config.consumer_key, 
						   consumer_secret=config.consumer_secret, 
						   access_token=config.access_token, 
						   access_token_secret=config.access_token_secret)
	return client

#Getting the tweets
def searchTweets(query):
	#Setting up the client to the variable client
	client = getClient()

	# The recent search endpoint returns Tweets from the last seven days that match a search query
	# In this project I will be getting tweets from 04/02/2022 to ...
	# from 10:00 to 20:00 every day to get as much tweets as possible.
	# I have got the Elevated Access of the Twitter API so I am allowed to get
	# maximum 100 results per request. So every hour I retrieve approximately 100 unique tweets
	# 10-11
	# 11-12
	# 12-13
	# 13-14
	# 14-15
	# 15-16
	# 16-17
	# 17-18
	# 18-19
	# 19-20
	
	tweets = client.search_recent_tweets(query=query, 
										 tweet_fields=['text','created_at','geo'], 
										 place_fields=['full_name','geo'],
										 expansions = 'geo.place_id',
										 start_time='2022-03-02T19:00:00-00:00',
										 end_time = '2022-03-02T20:00:00-00:00',
										 max_results=100)

		 

	tweet_data = tweets.data
	tweet_includes = tweets.includes

	#list of places from the includes object. Not every tweet is geo-tagged
	if tweet_includes:
		places = {p['id']: p for p in tweets.includes['places']}

	#open excel document
	wb = openpyxl.load_workbook('Twitter_Dataset2.xlsx')

	if not tweet_data is None and len(tweet_data) > 0:

		for tweet in tweet_data:
			excel_input_list = []

			tweet_id = tweet.id
			tweet_text = str(tweet.text)

			''' 
				In our response, we get the list of places from the includes object,
				and we match on the place_id to get the relevant geo information 
				associated with the Tweet
			'''

			if not tweet.geo is None:
				if places[tweet.geo['place_id']]:
					place = places[tweet.geo['place_id']]
					location = place.full_name

			date_created = (tweet.created_at.date())

			excel_input_list.append(tweet_id)
			excel_input_list.append(tweet_text)
			excel_input_list.append(date_created)

			if not tweet.geo is None:
				excel_input_list.append(location)

			try:
				ws = wb['Sheet1']
				ws.append(excel_input_list)
				wb.save('Twitter_Dataset2.xlsx')
			except:
				print("Appending file error")

	else:
		return []

tweets = searchTweets('("electric vehicles" OR #electricvehicles OR #EVs OR #EV) lang:en -is:retweet -is:reply')

