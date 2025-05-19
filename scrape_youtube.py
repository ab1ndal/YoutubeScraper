import os
from dotenv import load_dotenv
import csv
from googleapiclient.discovery import build
from datetime import datetime, timezone
import isodate
import pandas as pd

load_dotenv()   

API_KEY = os.getenv("YOUTUBE_API_KEY")

youtube = build('youtube', 'v3', developerKey=API_KEY)

def get_category_name(category_id):
    category_response = youtube.videoCategories().list(
        part='snippet',
        id=category_id
    ).execute()

    if category_response['items']:
        return category_response['items'][0]['snippet']['title']
    return 'Unknown'

def parse_duration(iso_duration):
    duration = isodate.parse_duration(iso_duration)
    return str(duration)

def calculate_engagement(likes, comments, views):
    try:
        return round((likes + comments) / views, 4)
    except (ZeroDivisionError, TypeError):
        return 0.0

def calculate_views_per_day(views, published_at):
    try:
        published_date = datetime.strptime(published_at, "%Y-%m-%dT%H:%M:%SZ")
        days_since = (datetime.now(timezone.utc) - published_date.replace(tzinfo=timezone.utc)).days or 1
        return round(views / days_since, 2)
    except Exception:
        return 0.0

def search_youtube(keyword, max_results=15):
    search_response = youtube.search().list(
        q=keyword,
        part='id,snippet',
        type='video',
        maxResults=max_results,
        order='viewCount',
        regionCode='US'
    ).execute()

    videos_data = []

    for item in search_response['items']:
        video_id = item['id']['videoId']
        channel_id = item['snippet']['channelId']

        # Video details
        video_details = youtube.videos().list(
            part='snippet,statistics,contentDetails',
            id=video_id
        ).execute()

        video_info = video_details['items'][0]
        snippet = video_info['snippet']
        stats = video_info['statistics']
        content = video_info['contentDetails']

        tags = snippet.get('tags', [])
        views = int(stats.get('viewCount', 0))
        likes = int(stats.get('likeCount', 0))
        comments = int(stats.get('commentCount', 0))
        published_at = snippet.get('publishedAt', '')
        duration = parse_duration(content.get('duration', 'PT0S'))
        category_name = get_category_name(snippet.get('categoryId', ''))

        engagement = calculate_engagement(likes, comments, views)
        views_per_day = calculate_views_per_day(views, published_at)

        # Channel info
        channel_details = youtube.channels().list(
            part='snippet,statistics',
            id=channel_id
        ).execute()

        channel_info = channel_details['items'][0]
        channel_name = channel_info['snippet']['title']
        subscribers = channel_info['statistics'].get('subscriberCount', 'Hidden')

        videos_data.append({
            'Channel': channel_name,
            'Subscribers': subscribers,
            'Video Title': snippet['title'],
            'Video Link': f"https://www.youtube.com/watch?v={video_id}",
            'Views': views,
            'Likes': likes,
            'Comments': comments,
            'Engagement Rate': engagement,
            'Views per Day': views_per_day,
            'Published At': published_at,
            'Duration': duration,
            'Category': category_name,
            'Hashtags': ', '.join([t for t in tags if t.startswith('#')]),
            'Description': snippet['description']
        })

    return videos_data

def save_to_excel(data, filename='youtube_analytics.xlsx'):
    if not data:
        print("No data to save.")
        return
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"Saved {len(data)} results to {filename}")

# 🔍 Run script
if __name__ == "__main__":
    keyword = "structure"  # You can change this to any keyword
    results = search_youtube(keyword)
    save_to_excel(results, 'youtube_analytics_{}.xlsx'.format(keyword))