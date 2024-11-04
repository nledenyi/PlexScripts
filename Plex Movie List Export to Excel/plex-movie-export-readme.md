# Plex Movie Library Exporter

A Python script that connects to your local Plex Media Server installation and generates an Excel spreadsheet containing details about your movie collection. This tool helps you maintain an offline catalog of your Plex movie library with key information such as video resolution, year, studio, and content rating.

## Features

- Connects to local Plex Media Server
- Extracts comprehensive movie information including:
  - Title
  - Video Resolution
  - Release Year
  - Studio
  - Content Rating
  - File Location
  - Container Format
- Exports data to a well-formatted Excel spreadsheet
- Automatically sorts movies alphabetically by title
- Handles missing data gracefully
- Provides progress feedback during execution

## Prerequisites

- Python 3.6 or higher
- Access to a Plex Media Server
- The following Python packages:
  - plexapi
  - pandas
  - openpyxl

## Installation

1. Clone this repository:
```bash
git clone https://github.com/PrimePoobah/plex-movie-exporter.git
cd plex-movie-exporter
```

2. Install required packages:
```bash
pip install plexapi pandas openpyxl
```

## Configuration

Before running the script, modify the `PLEX_URL` variable in the script to match your Plex server's address:

```python
PLEX_URL = '{Plex IP or URL}:32400'  # Replace with your Plex server address
```

If your Plex server requires authentication, you'll need to add your Plex token to the `connect_to_plex` function call:

```python
plex = connect_to_plex(PLEX_URL, token='your-plex-token-here')
```

To find your Plex token, follow the instructions in the [Plex documentation](https://support.plex.tv/articles/204059436-finding-an-authentication-token-x-plex-token/).

## Usage

Run the script from the command line:

```bash
python plex_movie_export.py
```

The script will:
1. Connect to your Plex server
2. Retrieve your movie library
3. Process all movie details
4. Create an Excel file named `plex_movies.xlsx` in the current directory

## Output

The script generates an Excel file (`plex_movies.xlsx`) with the following columns:
- Title: The movie's title
- Video Resolution: The resolution of the video (e.g., 4K, 1080p, 720p)
- Year: Release year of the movie
- Studio: Production studio
- ContentRating: Movie rating (e.g., PG, R, etc.)
- File: Full path to the movie file
- Container: File container format (e.g., mkv, mp4)

## Error Handling

The script includes basic error handling for:
- Failed server connections
- Missing media information
- Invalid or missing data fields

If any information is unavailable for a particular movie, the field will be marked as "Unknown" rather than failing.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## Possible Improvements

- Add command-line arguments for server address and authentication
- Include more movie metadata (genres, actors, etc.)
- Add filtering options for the movie list
- Support for multiple libraries
- Custom Excel formatting options
- Export to different formats (CSV, JSON, etc.)
- Add logging functionality
- Include movie poster extraction
- Add batch processing capabilities

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.

## Acknowledgments

- [PlexAPI](https://github.com/pkkid/python-plexapi) for providing the Python interface to Plex
- [Pandas](https://pandas.pydata.org/) for data handling and Excel export capabilities
- [Plex](https://www.plex.tv/) for their amazing media server platform

## Support

For support, please:
1. Check existing issues or create a new one
2. Provide detailed information about your setup and the error you're encountering
3. Include relevant logs and error messages
