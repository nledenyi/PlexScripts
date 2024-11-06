################################################################################
# Plex Media Export Script
# 
# Purpose: This script generates a comprehensive Excel report of a Plex Media Server's
# content, including both movies and TV shows. For TV shows, it verifies completion
# status against TVMaze's database.
#
# Features:
# - Exports movie details including resolution, year, and format
# - Tracks TV show completion status with color coding
# - Creates a professionally formatted Excel file with multiple worksheets
# - Includes progress reporting and error handling
################################################################################

# Import required external libraries
from plexapi.server import PlexServer   # For connecting to and querying Plex server
import pandas as pd                     # For data manipulation and initial Excel handling
from openpyxl import Workbook           # For creating and formatting Excel workbooks
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment  # For Excel styling
from openpyxl.utils import get_column_letter    # For Excel column references
import requests                         # For making HTTP requests to TVMaze API
from datetime import datetime           # For generating timestamps
import sys                              # For system operations (like exit)

# Configuration constants
# These values need to be replaced with actual Plex server details before running
PLEX_URL = 'http://{Plex_IP_or_URL}:32400'  # URL of your Plex server
PLEX_TOKEN = '{PlexToken}'                  # Your Plex authentication token
TVMAZE_API = 'https://api.tvmaze.com'       # Base URL for TVMaze API


# Define Excel cell fill colors for different completion states in TV Shows worksheet
GRAY_FILL = PatternFill(
    start_color='D3D3D3',   # Light gray for non-existent seasons
    end_color='D3D3D3',
    fill_type='solid'
)
GREEN_FILL = PatternFill(
    start_color='90EE90',   # Light green for complete seasons
    end_color='90EE90',
    fill_type='solid'
)
RED_FILL = PatternFill(
    start_color='FFB6B6',   # Light red for incomplete seasons
    end_color='FFB6B6',
    fill_type='solid'
)

LOW_RES_FILL = PatternFill(
    start_color='FFFFCC',   # Light yellow for low resolution videos
    end_color='FFFFCC',
    fill_type='solid'
)

FOURK_FILL = PatternFill(
    start_color='E3F4EA',   # Light green for 4K videos
    end_color='E3F4EA',
    fill_type='solid'
)

# Define Excel border styles for consistent cell formatting
THIN_BORDER = Border(
    left=Side(style='thin'),    # Standard border for all cells
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

HEADER_BORDER = Border(
    left=Side(style='thin'),    # Special border for header row with thick bottom
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thick')
)

def connect_to_plex():
    """
    Establishes connection to Plex server using configured URL and token.
    
    Returns:
        PlexServer: Connected Plex server instance
    
    Raises:
        SystemExit: If connection fails
    
    Notes:
        - Exits script with error message if connection fails
        - Uses global PLEX_URL and PLEX_TOKEN constants
    """
    try:
        # Attempt to create connection using provided credentials
        return PlexServer(PLEX_URL, PLEX_TOKEN)
    except Exception as e:
        # Log error and exit if connection fails
        print(f"Failed to connect to Plex server: {e}")
        sys.exit(1)

def get_movie_details(movies):
    """
    Extracts relevant details from movie objects in Plex library.
    
    Args:
        movies (list): List of Plex movie objects
    
    Returns:
        list: List of dictionaries containing movie information
              Each dictionary contains:
              - Title: Movie title
              - Video Resolution: Quality (4k, 1080p, etc.)
              - Year: Release year
              - Studio: Production studio
              - ContentRating: Rating (PG, R, etc.)
              - File: Full file path
              - Container: File format (mkv, mp4, etc.)
    """
    movie_list = []
    for movie in movies:
        # Get the first media item (typically there's only one)
        media = movie.media[0] if movie.media else None
        
        # Create dictionary with movie information
        # Use 'Unknown' as fallback for missing attributes
        movie_info = {
            'Title': movie.title,
            'Video Resolution': media.videoResolution if media else 'Unknown',
            'Year': movie.year,
            'Studio': movie.studio,
            'ContentRating': movie.contentRating,
            'File': media.parts[0].file if media and media.parts else 'Unknown',
            'Container': media.container if media else 'Unknown'
        }
        movie_list.append(movie_info)
    return movie_list

def get_tvmaze_show_info(show_name):
    """
    Retrieves show information from TVMaze API including season and episode counts.
    
    Args:
        show_name (str): Name of the TV show to search for
    
    Returns:
        dict: Dictionary containing:
             - total_seasons: Highest season number
             - seasons: Dict of season info including episode counts
             Returns None if show is not found or API request fails
    
    Notes:
        - Makes two API calls: one to search, one to get episodes
        - Aggregates episode counts by season
    """
    # First API call: Search for the show
    search_url = f"{TVMAZE_API}/search/shows"
    response = requests.get(search_url, params={'q': show_name})
    
    if response.status_code == 200 and response.json():
        # Extract show ID from search results
        show_id = response.json()[0]['show']['id']
        
        # Second API call: Get episode list
        episodes_url = f"{TVMAZE_API}/shows/{show_id}/episodes"
        episodes_response = requests.get(episodes_url)
        
        if episodes_response.status_code == 200:
            episodes = episodes_response.json()
            seasons = {}
            
            # Process each episode to build season information
            for episode in episodes:
                season_num = episode['season']
                # Initialize season if not seen before
                if season_num not in seasons:
                    seasons[season_num] = {'total_episodes': 0}
                seasons[season_num]['total_episodes'] += 1
            
            return {
                'total_seasons': max(seasons.keys()),
                'seasons': seasons
            }
    return None

def get_plex_show_info(plex, show):
    """
    Gets season and episode information for a show from Plex library.
    
    Args:
        plex (PlexServer): Connected Plex server instance
        show (PlexShow): Plex show object
    
    Returns:
        dict: Dictionary containing season information from Plex
             Keys are season numbers, values are dicts containing:
             - episodes_in_plex: Number of episodes in Plex
             - season_number: Season number
    """
    seasons = {}
    # Iterate through each season in the show
    for season in show.seasons():
        seasons[season.seasonNumber] = {
            'episodes_in_plex': len(season.episodes()),
            'season_number': season.seasonNumber
        }
    return seasons

def count_complete_seasons(show_data):
    """
    Counts how many seasons are complete in Plex compared to TVMaze data.
    
    Args:
        show_data (dict): Show information containing both Plex and TVMaze data
    
    Returns:
        int: Number of complete seasons
    
    Notes:
        - A season is complete when Plex episode count matches TVMaze
        - Only counts seasons that exist in TVMaze data
    """
    complete_seasons = 0
    # Check each season up to the highest season number
    for season_num in range(1, show_data['tvmaze_info']['total_seasons'] + 1):
        # Get episode counts from both sources
        plex_episodes = show_data['seasons'].get(season_num, {}).get('episodes_in_plex', 0)
        total_episodes = show_data['tvmaze_info']['seasons'].get(season_num, {}).get('total_episodes', 0)
        
        # Increment counter if all episodes are present
        if plex_episodes == total_episodes and total_episodes > 0:
            complete_seasons += 1
    return complete_seasons

def apply_borders_to_worksheet(ws, max_row, max_col):
    """
    Applies borders to all cells in the worksheet.
    
    Args:
        ws: Worksheet object
        max_row: Maximum row number
        max_col: Maximum column number
    
    Notes:
        - Header row (row 1) gets thick bottom border
        - All other cells get thin borders on all sides
    """
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            # Apply appropriate border style based on row
            if row == 1:
                cell.border = HEADER_BORDER
            else:
                cell.border = THIN_BORDER

def create_tv_shows_worksheet(wb, shows_data, max_seasons):
    """
    Creates and populates the TV Shows worksheet in the workbook.
    
    Args:
        wb (Workbook): Excel workbook object
        shows_data (list): List of dictionaries containing show information
        max_seasons (int): Highest number of seasons among all shows
    
    Notes:
        - Creates headers for show title, completion status, and each season
        - Applies formatting including:
            * Bold headers
            * Centered alignment (except show titles)
            * Color coding for completion status
            * Borders on all cells
            * Frozen header row
    """
    ws = wb.create_sheet("TV Shows")
    
    # Create headers
    headers = ["Show Title", "Complete Series"]
    headers.extend([f"Season {i}" for i in range(1, max_seasons + 1)])
    
    # Write headers to worksheet and make them bold
    bold_font = Font(bold=True)
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal='center')
    
    # Freeze the first row
    ws.freeze_panes = 'A2'
    
    # Process each show
    row = 2
    for show in shows_data:
        # Write show title (left-aligned)
        title_cell = ws.cell(row=row, column=1, value=show['title'])
        title_cell.alignment = Alignment(horizontal='left')
        
        # Calculate and write completion status
        complete_seasons = count_complete_seasons(show)
        total_seasons = show['tvmaze_info']['total_seasons']
        
        # Format the completion status cell
        series_cell = ws.cell(row=row, column=2)
        series_cell.value = f"{complete_seasons}/{total_seasons}"
        series_cell.alignment = Alignment(horizontal='center')
        
        # Color code the series completion cell
        if complete_seasons == total_seasons:
            series_cell.fill = GREEN_FILL  # All seasons complete
        elif complete_seasons > 0:
            series_cell.fill = RED_FILL    # Partially complete
        
        # Process each season
        for season_num in range(1, max_seasons + 1):
            cell = ws.cell(row=row, column=season_num + 2)
            cell.alignment = Alignment(horizontal='center')
            
            if season_num <= total_seasons:
                plex_episodes = show['seasons'].get(season_num, {}).get('episodes_in_plex', 0)
                total_episodes = show['tvmaze_info']['seasons'].get(season_num, {}).get('total_episodes', 0)
                
                if total_episodes > 0:
                    cell.value = f"{plex_episodes}/{total_episodes}"
                    # Color code based on completion
                    if plex_episodes == total_episodes:
                        cell.fill = GREEN_FILL  # Complete season
                    elif plex_episodes > 0:
                        cell.fill = RED_FILL    # Partial season
            else:
                cell.fill = GRAY_FILL  # Season doesn't exist
        
        row += 1
    
    # Apply borders to all cells
    apply_borders_to_worksheet(ws, row - 1, len(headers))
    
    # Adjust column widths for better readability
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width

def create_movies_worksheet(wb, movie_list):
    """
    Creates and populates the Movies worksheet in the workbook.
    
    Args:
        wb (Workbook): Excel workbook object
        movie_list (list): List of dictionaries containing movie information
    
    Notes:
        - Converts movie list to DataFrame for easier handling
        - Applies formatting including:
            * Bold headers
            * Centered alignment (except title and file path)
            * Borders on all cells
            * Frozen header row
            * Auto-adjusted column widths
            * Highlights low resolution videos (sd, 480, 720) in light yellow
            * Highlights 4K videos in light green
    """
    # Convert movie list to DataFrame for easier handling
    df = pd.DataFrame(movie_list)
    df = df.sort_values('Title')  # Sort alphabetically by title
    
    # Create the Movies worksheet
    ws = wb.create_sheet("Movies")
    
    # Write headers and make them bold
    bold_font = Font(bold=True)
    for col, header in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal='center')
    
    # Freeze the first row
    ws.freeze_panes = 'A2'
    
    # Write movie data with appropriate alignment and highlighting
    for row in range(len(df)):
        # Check resolution for row highlighting
        resolution = str(df.iloc[row]['Video Resolution']).lower()
        
        # Determine highlighting based on resolution
        if resolution in ['sd', '480', '720']:
            row_fill = LOW_RES_FILL
        elif resolution in ['4k', 'uhd']:  # Including 'uhd' as some systems might use this
            row_fill = FOURK_FILL
        else:
            row_fill = None
        
        for col in range(len(df.columns)):
            cell = ws.cell(row=row+2, column=col+1, value=str(df.iloc[row, col]))
            
            # Center all columns except Title and File path
            if col not in [0, 4]:  # Assuming Title is column 0 and File is column 4
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.alignment = Alignment(horizontal='left')
            
            # Apply highlight if needed
            if row_fill:
                cell.fill = row_fill
    
    # Apply borders to all cells
    apply_borders_to_worksheet(ws, len(df) + 1, len(df.columns))
    
    # Adjust column widths for better readability
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width
    """
    Creates and populates the Movies worksheet in the workbook.
    
    Args:
        wb (Workbook): Excel workbook object
        movie_list (list): List of dictionaries containing movie information
    
    Notes:
        - Converts movie list to DataFrame for easier handling
        - Applies formatting including:
            * Bold headers
            * Centered alignment (except title and file path)
            * Borders on all cells
            * Frozen header row
            * Auto-adjusted column widths
            * Highlights low resolution videos (sd, 480, 720) in light yellow
    """
    # Convert movie list to DataFrame for easier handling
    df = pd.DataFrame(movie_list)
    df = df.sort_values('Title')  # Sort alphabetically by title
    
    # Create the Movies worksheet
    ws = wb.create_sheet("Movies")
    
    # Write headers and make them bold
    bold_font = Font(bold=True)
    for col, header in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal='center')
    
    # Freeze the first row
    ws.freeze_panes = 'A2'
    
    # Write movie data with appropriate alignment and highlighting
    for row in range(len(df)):
        # Check resolution for row highlighting
        resolution = str(df.iloc[row]['Video Resolution']).lower()
        should_highlight = resolution in ['sd', '480', '720']
        
        for col in range(len(df.columns)):
            cell = ws.cell(row=row+2, column=col+1, value=str(df.iloc[row, col]))
            
            # Center all columns except Title and File path
            if col not in [0, 4]:  # Assuming Title is column 0 and File is column 4
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.alignment = Alignment(horizontal='left')
            
            # Apply highlight if needed
            if should_highlight:
                cell.fill = LOW_RES_FILL
    
    # Apply borders to all cells
    apply_borders_to_worksheet(ws, len(df) + 1, len(df.columns))
    
    # Adjust column widths for better readability
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width
    """
    Creates and populates the Movies worksheet in the workbook.
    
    Args:
        wb (Workbook): Excel workbook object
        movie_list (list): List of dictionaries containing movie information
    
    Notes:
        - Converts movie list to DataFrame for easier handling
        - Applies formatting including:
            * Bold headers
            * Centered alignment (except title and file path)
            * Borders on all cells
            * Frozen header row
            * Auto-adjusted column widths
    """
    # Convert movie list to DataFrame for easier handling
    df = pd.DataFrame(movie_list)
    df = df.sort_values('Title')  # Sort alphabetically by title
    
    # Create the Movies worksheet
    ws = wb.create_sheet("Movies")
    
    # Write headers and make them bold
    bold_font = Font(bold=True)
    for col, header in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal='center')
    
    # Freeze the first row
    ws.freeze_panes = 'A2'
    
    # Write movie data with appropriate alignment
    for row in range(len(df)):
        for col in range(len(df.columns)):
            cell = ws.cell(row=row+2, column=col+1, value=str(df.iloc[row, col]))
            # Center all columns except Title and File path
            if col not in [0, 4]:  # Assuming Title is column 0 and File is column 4
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.alignment = Alignment(horizontal='left')
    
    # Apply borders to all cells
    apply_borders_to_worksheet(ws, len(df) + 1, len(df.columns))
    
    # Adjust column widths for better readability
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width

def main():
    """
    Main function to coordinate the media export process.
    
    Process:
    1. Connects to Plex server
    2. Processes both movies and TV shows
    3. Generates a combined Excel report
    4. Saves the report with a timestamp
    
    Notes:
        - Provides progress updates during processing
        - Handles errors gracefully
        - Creates a timestamped output file
    """
    # Connect to Plex server
    print("Connecting to Plex server...")
    plex = connect_to_plex()
    
    # Process Movies
    print("Processing movies...")
    movies_section = plex.library.section('Movies')
    all_movies = movies_section.all()
    movie_list = get_movie_details(all_movies)
    
    # Process TV Shows
    print("Processing TV shows...")
    tv_section = plex.library.section('TV Shows')
    shows_data = []
    max_seasons = 0
    
    # Process each TV show
    for show in tv_section.all():
        print(f"Processing: {show.title}")
        # Get TVMaze information
        tvmaze_info = get_tvmaze_show_info(show.title)
        if not tvmaze_info:
            print(f"Could not find TVMaze info for: {show.title}")
            continue
            
        # Track maximum number of seasons
        if tvmaze_info['total_seasons'] > max_seasons:
            max_seasons = tvmaze_info['total_seasons']
            
        # Get Plex information
        plex_seasons = get_plex_show_info(plex, show)
        shows_data.append({
            'title': show.title,
            'seasons': plex_seasons,
            'tvmaze_info': tvmaze_info
        })
    
    # Create Excel workbook
    print("Creating Excel report...")
    wb = Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Create both worksheets
    create_movies_worksheet(wb, movie_list)
    create_tv_shows_worksheet(wb, shows_data, max_seasons)
    
    # Save the workbook with timestamp
    timestamp = datetime.now().strftime('%Y%m%d')
    filename = f"PlexMediaExport_{timestamp}.xlsx"
    wb.save(filename)
    
    # Print completion message
    print(f"Export complete! Found {len(movie_list)} movies and {len(shows_data)} TV shows.")
    print(f"Report saved as: {filename}")

# Script entry point
if __name__ == "__main__":
    main()