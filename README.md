# Plex

This repository provides Python scripts designed to export content from a Plex Media Server into Excel files. The exporter allows users to retrieve detailed information about movies, TV shows, or both, complete with color-coded status indicators for efficient organization and analysis.

## Scripts

### Plex Movie Exporter
A Python script that connects to your local Plex Media Server installation and generates an Excel spreadsheet containing details about your movie collection. This tool helps you maintain an offline catalog of your Plex movie library with key information such as video resolution, year, studio, and content rating.

##### Features
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

### Plex TV Show Exporter
A Python script that audits your Plex TV Show library against TVMaze data to create a detailed Excel report showing which shows and seasons are complete or missing episodes.

##### Features
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

### Plex Media Export
A Python utility that generates a detailed Excel report of your Plex Media Server library. The tool provides an easy way to keep track of your media collection, highlighting video quality and tracking TV show completeness through TVMaze integration.

##### Features
- This script creates a comprehensive Excel spreadsheet with two worksheets:
  - Movies: Displays all movies with resolution-based highlighting
  - TV Shows: Shows episode completion status verified against TVMaze
- Complete movie library inventory
- Resolution-based row highlighting:
  - 4K content (Light Green)
  - SD/480p/720p content (Yellow)
  - 1080p content (No highlight)
- Video format and file path information
- TV Show Tracking
- Series completion overview
- Season-by-season episode verification
- Color-coded status indicators
- TVMaze database integration

## Contributing

We welcome contributions to the Plex project! If you have any suggestions, bug reports, or feature requests, please open an issue or submit a pull request.

## License

This project is licensed under the GPL-3.0 License. See the [LICENSE](https://github.com/PrimePoobah/Plex) file for more details.
