from datetime import datetime, timedelta
import GetOldTweets3 as got
import xlsxwriter


def get_tweets():

    # set query search criteria with hashtag and time window
    tweets_criteria = got.manager.TweetCriteria().setQuerySearch("#ognunoèperfetto") \
        .setSince("2019-12-23") \
        .setUntil("2019-12-25")

    # download the tweets according to the search criterias
    tweets = got.manager.TweetManager().getTweets(tweets_criteria)

    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('tweets24CEST.xlsx')
    worksheet = workbook.add_worksheet()

    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 20)

    worksheet.write(0, 0, "DateTime")
    worksheet.write(0, 1, "Tweet")

    # iterate through tweets
    # for each tweet, extract the date and the tweet
    for i, tweet in enumerate(tweets, 1):

        # extract the date
        d = datetime.strptime(str(tweet.date), '%Y-%m-%d %H:%M:%S+00:00').strftime('%Y/%m/%d %H:%M')
        datetime_obj = datetime.strptime(d, '%Y/%m/%d %H:%M')

        # convert the date to UTC format
        datetime_obj = datetime_obj + timedelta(hours=2)
        d = datetime.strptime(str(datetime_obj), '%Y-%m-%d %H:%M:%S').strftime('%Y/%m/%d %H:%M')

        # write the entry to the Excel File and write to the console.
        print("{}. {}-->{}".format(i, str(d), tweet.text))
        worksheet.write(i, 0, str(d))
        worksheet.write(i, 1, tweet.text)

    workbook.close()


get_tweets()