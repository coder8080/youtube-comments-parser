import xlsxwriter
import shutil
from googleapiclient.discovery import build
import os

api_key = "<yout_api_key>"


def get_resource():
    return build('youtube', 'v3', developerKey=api_key)


def get_number_input(text, error_text='Type a number:'):
    print(text)
    i = input()
    while not i.isdigit():
        print(error_text)
        i = input()
    return int(i)


def main():
    resource = get_resource()

    # Get user input
    print('Insert channel id (specified in the url at the main channel page):')
    channel_id = input()
    c = get_number_input(
        'How many videos do you want to parse? (insert a number):')
    k = get_number_input(
        'How many comment per video do you want to parse? (insert a number):')

    # Get channel info
    channel_info = resource.channels().list(
        id=channel_id,
        part="contentDetails"
    ).execute()
    video_group_id = channel_info['items'][0]['contentDetails']['relatedPlaylists']['uploads']

    # Get videos from channel
    videos = resource.playlistItems().list(
        playlistId=video_group_id, part="snippet", maxResults=c).execute()['items']

    # Loop through all videos
    for i, video in enumerate(videos):
        # Get video info
        video_id = video['snippet']['resourceId']['videoId']
        video_title = video['snippet']['title']
        video_comments = resource.commentThreads().list(
            part="snippet",
            videoId=video_id,
            maxResults=k,
            order="relevance").execute()['items']

        # Create a directory for xlsx file
        path = f'data/{i + 1}_{video_title.replace("/", "_")}'
        if os.path.exists(path):
            shutil.rmtree(path)
        os.mkdir(path)

        # Prepare xlsx file
        workbook = xlsxwriter.Workbook(f'{path}/comments.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.write(0, 0, 'Автор')
        worksheet.write(0, 1, 'Текст')

        # Loop through comments and write them into the file
        for i, comment in enumerate(video_comments):
            author = comment['snippet']['topLevelComment']['snippet']['authorDisplayName']
            text = comment['snippet']['topLevelComment']['snippet']['textOriginal']
            worksheet.write(i + 1, 0, author)
            worksheet.write(i + 1, 1, text)
        workbook.close()


main()
