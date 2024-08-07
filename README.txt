# Asiakastieto Board Membership Scraper

This project is a web scraper designed to extract due diligence information from asiakastieto.fi. It retrieves information about an individual's board memberships and their fellow board members.

## Features

- Scrapes board membership information for specified individuals
- Retrieves details of other board members for each company
- Supports multiple person searches
- Generates a report file for each person searched
- Handles both companies and foundations ("säätiöt")

## Prerequisites

- Python 3.6+
- Google Chrome browser
- ChromeDriver (compatible with your Chrome version)

## Installation

1. Clone this repository:

git clone https://github.com/yourusername/asiakastieto-scraper.git
cd asiakastieto-scraper

2. Install required Python packages:
pip install -r requirements.txt

3. Ensure you have the correct ChromeDriver for your operating system and Chrome version in the `driver` folder.

4. On Linux and macOS, make the ChromeDriver executable:
chmod +x driver/linux_chromedriver
chmod +x driver/mac_chromedriver

## Usage

1. Run the script:
python scrape.py

2. Enter the name of the person you want to search for in the GUI.
3. For multiple searches, separate names by comma: "John Doe, Jane Smith".
4. Click "Search" or press Enter to start the search.
5. The results will be saved in the `dd` folder as .docx files.

## Troubleshooting

- If the program doesn't work, try running it again. Occasional failures can occur due to network issues or changes in the website structure.
- Ensure your Chrome browser is up to date.
- If you're using a newer Chrome version, you may need to download a compatible ChromeDriver from [ChromeDriver Downloads](https://chromedriver.chromium.org/downloads).
- For multiple searches, try searching for people individually if batch search fails.

## Known Issues

- The scraper may occasionally fail to retrieve information for certain companies or foundations. These will be listed at the end of the search process.

## Contributing

Contributions, issues, and feature requests are welcome. Feel free to check [issues page](https://github.com/yourusername/asiakastieto-scraper/issues) if you want to contribute.

## License

[MIT](https://choosealicense.com/licenses/mit/)

## Contact

If you have any questions or encounter issues, please contact:
Your Name - your.email@example.com

Project Link: [https://github.com/yourusername/asiakastieto-scraper](https://github.com/yourusername/asiakastieto-scraper)