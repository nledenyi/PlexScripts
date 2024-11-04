# Import required libraries
from plexapi.server import PlexServer  # Library for interacting with Plex Media Server
import pandas as pd                    # Library for data manipulation and Excel export
from urllib.parse import urlparse      # Library for parsing URLs
import sys                            # Library for system-level operations

# Configuration settings for the script
PLEX_URL = 'http://{Plex_IP_or_URL}}:32400'  # URL of your Plex server
PLEX_TOKEN = '{PlexToken}'  # Your Plex authentication token

def connect_to_plex(PLEX_URL, PLEX_TOKEN):
    """
    Establishes connection to Plex server
    Args:
        PLEX_URL (str): URL of the Plex server
        PLEX_TOKEN (str, optional): Authentication token for Plex
    Returns:
        PlexServer: Connected Plex server instance or exits on failure
    """
    try:
        # Attempt to create a connection to the Plex server
        return PlexServer(PLEX_URL, PLEX_TOKEN)
    except Exception as e:
        # If connection fails, print error message and exit
        print(f"Failed to connect to Plex server: {e}")
        sys.exit(1)

def get_movie_details(movies):
    """
    Extracts relevant details from movie objects
    Args:
        movies (list): List of Plex movie objects
    Returns:
        list: List of dictionaries containing movie information
    """
    # Initialize empty list to store movie information
    movie_list = []
    
    # Iterate through each movie in the library
    for movie in movies:
        # Get the first media item (typically there's only one)
        # Media contains technical details like resolution and container
        media = movie.media[0] if movie.media else None
        
        # Create a dictionary with movie information
        movie_info = {
            'Title': movie.title,                    # Movie title
            'Video Resolution': media.videoResolution if media else 'Unknown',  # Resolution (4k, 1080p, etc.)
            'Year': movie.year,                      # Release year
            'Studio': movie.studio,                  # Production studio
            'ContentRating': movie.contentRating,    # Rating (PG, R, etc.)
            # Full path to the movie file
            'File': media.parts[0].file if media and media.parts else 'Unknown',
            'Container': media.container if media else 'Unknown'  # File container (mkv, mp4, etc.)
        }
        
        # Add this movie's information to our list
        movie_list.append(movie_info)
    
    # Return the complete list of movie information
    return movie_list

def main():
    
    # Inform user that connection is being attempted
    print("Connecting to Plex server...")
    # Establish connection to the Plex server
    plex = connect_to_plex(PLEX_URL, PLEX_TOKEN)
    
    # Inform user that movie fetching is starting
    print("Fetching movie library...")
    # Get access to the Movies section of your Plex library
    movies_section = plex.library.section('Movies')
    # Fetch all movies from the library
    all_movies = movies_section.all()
    
    # Inform user that processing is starting
    print("Processing movie details...")
    # Extract details from all movies
    movie_list = get_movie_details(all_movies)
    
    # Convert the list of movie dictionaries to a pandas DataFrame
    # DataFrame provides easy data manipulation and export capabilities
    df = pd.DataFrame(movie_list)
    
    # Sort the DataFrame alphabetically by movie title
    df = df.sort_values('Title')
    
    # Define the output file name
    output_file = 'plex_movies.xlsx'
    # Inform user that export is starting
    print(f"Exporting to {output_file}...")
    # Export the DataFrame to an Excel file
    # index=False prevents the DataFrame index from being included
    df.to_excel(output_file, index=False, sheet_name='Movies')
    
    # Inform user that the process is complete and show number of movies processed
    print(f"Export complete! Found {len(movie_list)} movies.")

# This is the standard Python idiom for a script that can be run directly
# or imported as a module
if __name__ == "__main__":
    # Execute the main function when the script is run directly
    main()