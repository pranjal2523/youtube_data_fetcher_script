import openpyxl
import os
from dotenv import load_dotenv
from googleapiclient.discovery import build


load_dotenv()  # Loads environment variables from a .env file

# Replace with your own YouTube Data API key
API_KEY = os.environ.get(
    'API_KEY'
)

# Initialize the YouTube API client
youtube = build('youtube', 'v3', developerKey=API_KEY)


def get_channel_id_by_username(username):
    """
    Fetches the channel ID for a given username (or channel handle).
    
    Args:
        username (str): The YouTube channel handle.

    Returns:
        str or None: The channel ID if found, otherwise None.
    """
    try:
        request = youtube.channels().list(
            part='id',
            forHandle=username
        )
        response = request.execute()

        # Extract the channel ID
        items = response.get('items')
        if items:
            return items[0]['id']
        else:
            print(f"No channel found with title: {username}")
            return None
    except Exception as e:
        print(f"Error occurred while fetching channel ID: {e}")
        return None


def get_videos_by_channel_id(channel_id):
    """
    Fetches details of videos from a channel given its channel ID.

    Args:
        channel_id (str): The YouTube channel ID.

    Returns:
        list: A list of dictionaries containing video details.
    """
    videos = []
    next_page_token = None

    # Fetch videos from the channel
    while True:
        try:
            request = youtube.search().list(
                part='snippet',
                channelId=channel_id,
                maxResults=50,
                order='date',
                type='video',
                pageToken=next_page_token
            )
            response = request.execute()

            # Collect video details
            for item in response['items']:
                video_id = item['id']['videoId']
                
                # Fetch more detailed video information
                video_details = get_video_details(video_id)
                if video_details:
                    videos.append(video_details)

            # Check if there's another page of results
            next_page_token = response.get('nextPageToken')
            if not next_page_token:
                break
        except Exception as e:
            print(f"Error occurred while fetching videos: {e}")
            break

    return videos


def get_video_details(video_id):
    """
    Fetches detailed information about a specific video.

    Args:
        video_id (str): The YouTube video ID.

    Returns:
        dict or None: A dictionary containing video details, or None if not found.
    """
    try:
        request = youtube.videos().list(
            part='snippet, contentDetails, statistics',
            id=video_id
        )
        response = request.execute()
        if not response['items']:
            return None

        item = response['items'][0]
        snippet = item['snippet']
        content_details = item['contentDetails']
        statistics = item['statistics']
        thumbnails = snippet['thumbnails']

        return {
            'video_id': video_id,
            'title': snippet['title'],
            'description': snippet['description'],
            'published_at': snippet['publishedAt'],
            'duration': content_details['duration'],
            'view_count': statistics.get('viewCount'),
            'like_count': statistics.get('likeCount'),
            'comment_count': statistics.get('commentCount'),
            'default_thumbnail': thumbnails['default']['url'],
            'medium_thumbnail': thumbnails['medium']['url'],
            'high_thumbnail': thumbnails['high']['url']
        }
    except Exception as e:
        print(f"Error occurred while fetching video details for {video_id}: {e}")
        return None


def save_videos_to_excel(videos, filename="youtube_videos.xlsx"):
    """
    Saves the data from the videos list to an Excel file.

    Args:
        videos (list): A list of dictionaries containing video details.
        filename (str): The name of the Excel file to save data to.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Videos"

    # Set headers for the first row
    headers = [
        "Video ID", "Title", "Description", "Published At", "Duration",
        "View Count", "Like Count", "Comment Count", "Default Thumbnail",
        "Medium Thumbnail", "High Thumbnail"
    ]
    ws.append(headers)

    # Add video details to each row
    for video in videos:
        row_data = [
            video["video_id"], video["title"], video["description"], video["published_at"],
            video["duration"], video.get("view_count"), video.get("like_count"), video.get("comment_count"),
            video["default_thumbnail"], video["medium_thumbnail"], video["high_thumbnail"]
        ]
        ws.append(row_data)

    # Save the workbook
    wb.save(filename)
    print(f"Excel file saved successfully: {filename}")


def get_comments_by_video_id(video_id, max_comments=100):
    """
    Fetches the latest comments (and their replies) for a video.

    Args:
        video_id (str): The YouTube video ID.
        max_comments (int): The maximum number of comments to fetch.

    Returns:
        list: A list of dictionaries containing comment details.
    """
    comments = []
    next_page_token = None
    comment_count = 0

    # Fetch comments from the video
    while comment_count < max_comments:
        try:
            request = youtube.commentThreads().list(
                part='snippet',
                videoId=video_id,
                textFormat='plainText',
                maxResults=100,
                pageToken=next_page_token
            )
            response = request.execute()

            # Process comments and replies
            for item in response['items']:
                comment = item['snippet']['topLevelComment']['snippet']
                comment_id = item['id']
                comments.append({
                    'video_id': video_id,
                    'comment_id': comment_id,
                    'comment_text': comment['textDisplay'],
                    'author_name': comment['authorDisplayName'],
                    'published_at': comment['publishedAt'],
                    'like_count': comment['likeCount'],
                    'reply_to': None
                })

                # Handle replies, if any
                if 'replies' in item:
                    for reply in item['replies']['comments']:
                        reply_snippet = reply['snippet']
                        comments.append({
                            'video_id': video_id,
                            'comment_id': reply['id'],
                            'comment_text': reply_snippet['textDisplay'],
                            'author_name': reply_snippet['authorDisplayName'],
                            'published_at': reply_snippet['publishedAt'],
                            'like_count': reply_snippet['likeCount'],
                            'reply_to': comment_id
                        })

            comment_count += len(response['items'])
            next_page_token = response.get('nextPageToken')
            if not next_page_token:
                break
        except Exception as e:
            print(f"Error occurred while fetching comments: {e}")
            break

    return comments


def export_comments_to_excel(comments, filename="youtube_comments.xlsx"):
    """
    Exports the comments data to an Excel file.

    Args:
        comments (list): A list of dictionaries containing comment details.
        filename (str): The name of the Excel file to save data to.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comments"

    # Set headers for the first row
    headers = [
        "Video ID", "Comment ID", "Comment Text", "Author Name",
        "Published At", "Like Count", "Reply To"
    ]
    ws.append(headers)

    # Add comment details to each row
    for comment in comments:
        row_data = [
            comment["video_id"], comment["comment_id"], comment["comment_text"], comment["author_name"],
            comment["published_at"], comment["like_count"], comment["reply_to"]
        ]
        ws.append(row_data)

    # Save the workbook
    wb.save(filename)
    print(f"Excel file saved successfully: {filename}")


def main():
    """
    Main function to fetch videos and comments for a given YouTube channel.
    """
    username = input("Enter YouTube channel username: ")
    try:
        max_comments = int(input("Enter the number of comments you want to fetch: "))
    except ValueError:
        print("Invalid value for comment count!")
        return
    try:
        username = username.split('/')[-1]
    except IndexError:
        print("Invalid URL. Please check & try again!")
        return 

    channel_id = get_channel_id_by_username(username)
    if not channel_id:
        return

    print(f"Fetching videos for channel ID: {channel_id}")
    videos = get_videos_by_channel_id(channel_id)
    save_videos_to_excel(videos, f'videos_data_{username}.xlsx')

    all_comments = []
    for video in videos:
        if len(all_comments) >= max_comments:
            break
        video_id = video['video_id']
        # print(f"Fetching comments for video ID: {video_id}")
        comments = get_comments_by_video_id(video_id, max_comments=max_comments - len(all_comments))
        all_comments.extend(comments)

    export_comments_to_excel(all_comments, filename=f'comments_data_of_{username}.xlsx')


if __name__ == '__main__':
    main()
