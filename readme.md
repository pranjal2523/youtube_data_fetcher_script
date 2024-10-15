# YouTube Channel Video and Comments Data Fetcher

This Python script fetches video data and comments from a YouTube channel using the YouTube Data API v3. It retrieves basic details for each video and the latest 100 comments along with replies. The output is saved in two separate Excel files: one for video data and another for comments data.

## Features

- Fetches channel videos using the channel's handle.
- Retrieves detailed video information such as title, description, duration, view count, like count, comment count, and thumbnails.
- Fetches comments and replies for each video, including author information, likes, and timestamps.
- Saves the data to Excel files with user-specified filenames.

## Prerequisites

Before using this script, ensure you have:

1. A Google Cloud Platform account with access to the YouTube Data API v3.
2. Python 3.x installed on your system.
3. An API key for YouTube Data API v3.

## Setup

### 1. Get YouTube Data API Key

To use the YouTube Data API, you'll need to create a project on the Google Cloud Platform and enable the YouTube Data API v3. Follow these steps:

1. Visit the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project.
3. Enable the YouTube Data API v3 for the project.
4. Create an API key and note it down.

### 2. Install Required Python Libraries

Install the required libraries using `pip`:

```bash
    pip install google-api-python-client openpyxl python-dotenv
```
### 3. Create a .env file in the root directory to store your environment variables (API keys, etc.).
    ```
    API_KEY = "API_KEY_VALUE"
    ```
## Usage
1. Clone or download this repository.
2. Set up the API_KEY environment variable as mentioned above.
3. Run the script:
    ```
    python youtube_data_fetch.py
    ```

## Script Input
    - Enter the YouTube channel username (handle) when prompted.
    - Enter the number of comments you want to fetch (up to 100 comments).

## Output
    - Two Excel files will be created:
    - Videos_data_<username>.xlsx: Contains video data for the specified channel.
    - comments_data_of_<username>.xlsx: Contains comment data for the videos.


## Code Explanation

The script is divided into several functions:

1. get_channel_id_by_username(username)
    #### Fetches the channel ID for a given YouTube channel username (handle). It uses the YouTube Data API's channels().list method to get the channel ID.

2. get_videos_by_channel_id(channel_id)
    #### Fetches a list of videos from the specified channel. It retrieves details such as:

    -  Video ID
    -  Title
    -  Description
    -  Published date
    -  Duration
    -  View count, like count, and comment count
    -  Thumbnails (default, medium, high resolution)

3. save_videos_to_excel(videos, filename="youtube_videos.xlsx")
    #### Saves the fetched video details to an Excel file using the openpyxl library.

4. get_comments_by_video_id(comments, video_id, max_comments=100)
    #### Fetches the latest comments and replies for a given video ID, including details such as:

    -  Comment ID
    -  Author name
    -  Comment text
    -  Published date
    -  Like count
    -  Reply information (if applicable)

5. export_comments_to_excel(comments, filename="youtube_comments.xlsx")
    - Exports the comments data to an Excel file.

6. main()
    - The main function orchestrates the script's flow:

### Takes user input for the channel username and number of comments.
### Fetches channel videos and saves them to an Excel file.
### Fetches comments for the videos and saves them to another Excel file.


## Error Handling
    -  The script checks if a valid API key is set; otherwise, it raises a ValueError.
    -  If no channel is found for the provided username, the script notifies the user.
    -  If no comments are found, the script continues without raising an error.