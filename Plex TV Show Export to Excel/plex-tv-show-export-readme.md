# Plex TV Show Audit Tool

A Python script that audits your Plex TV Show library against TVMaze data to create a detailed Excel report showing which shows and seasons are complete or missing episodes.

## Features

- Automatically scans your entire Plex TV Show library
- Cross-references each show with TVMaze to get official episode counts
- Generates a color-coded Excel report showing:
  - Complete series status
  - Episode counts for each season
  - Missing episodes highlighted in red
  - Complete seasons highlighted in green
  - Non-existent seasons grayed out
- Dynamic column generation based on the maximum number of seasons found
- Detailed progress logging during execution

## Screenshots

[Add screenshots of the Excel report here]

## Prerequisites

- Python 3.6 or higher
- A Plex Media Server installation
- A Plex authentication token
- Internet access (for TVMaze API)

## Required Python Packages

```bash
pip install plexapi requests pandas openpyxl
```

## Configuration

Before running the script, you need to configure the following variables in the script:

```python
PLEX_URL = 'http://{Plex_IP_or_URL}:32400'  # Your Plex server URL
PLEX_TOKEN = 'YOUR_PLEX_TOKEN'              # Your Plex authentication token
```

### Getting Your Plex Token

1. Sign in to Plex web app
2. View any video/audio file
3. Click the three dots menu (...)
4. Click "Get Info"
5. Click "View XML"
6. In the URL, look for "X-Plex-Token="
7. Copy the token that follows

## Usage

1. Clone the repository:
```bash
git clone https://github.com/yourusername/plex-tv-audit.git
cd plex-tv-audit
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Update the configuration variables in the script

4. Run the script:
```bash
python plex_tv_audit.py
```

The script will generate an Excel file with the current date in the filename (e.g., `plex_tv_shows_20241103.xlsx`).

## Excel Report Format

The generated Excel report includes:

- **Show Title**: Name of the TV show
- **Complete Series**: Number of complete seasons vs total seasons
- **Season Columns**: One column per season up to the highest season number found
- **Color Coding**:
  - 🟩 Green: Complete season/series
  - 🟥 Red: Incomplete season/series (missing episodes)
  - ⬜ Gray: Season doesn't exist for this show

Each season cell shows: `episodes in Plex / total episodes`

## Error Handling

The script handles several common error cases:
- Unable to connect to Plex server
- Show not found in TVMaze
- Missing or incomplete season data
- API connection issues

Errors are logged to the console during execution.

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b new-feature`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin new-feature`
5. Submit a pull request

## Known Issues

- TVMaze API has rate limiting that may affect large libraries
- Some shows may not match exactly between Plex and TVMaze due to naming differences
- Special episodes might not be counted correctly

## Future Improvements

- [ ] Add support for multiple Plex libraries
- [ ] Add support for movies
- [ ] Add HTML report option
- [ ] Add command-line arguments for configuration
- [ ] Add support for other metadata providers
- [ ] Add progress bar for large libraries
- [ ] Add option to export missing episodes list

## License

This project is licensed under the GPL-3.0 License. See the [LICENSE.md](LICENSE.md) file for details.

## Acknowledgments

- [Plex](https://www.plex.tv/) for their media server
- [TVMaze](https://www.tvmaze.com/) for their comprehensive TV show database
- [PlexAPI](https://python-plexapi.readthedocs.io/) for the Python Plex interface

## Support

For support, please open an issue in the GitHub repository or contact [your contact information].

## Author

Your Name
- GitHub: [@yourusername](https://github.com/yourusername)
- Email: your.email@example.com

---
**Note**: This project is not affiliated with or endorsed by Plex or TVMaze.
