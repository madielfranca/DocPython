# Scraping New York Times website

This project uses Python and Selenium library to scrape the New York Times website for news articles. The program allows the user to select a search phrase, category, and date range for the articles they want to retrieve. The data is extracted and saved to a CSV file.
Prerequisites

This program requires Python and the following libraries:

    RPA Framework
    Selenium
    Robocorp

The program can be run on any operating system that supports Python.

#Getting Started

To use this program, follow these steps:

    Clone the repository
    Install the prerequisites using pip or any other package manager
    Check the setup.py file and set your variables
    Run the main.py file with Python

How to use

To use the program, open the main.py file and set the following variables in the setup.py file:

    URL: The URL for the New York Times website
    SEARCH_PHRASE: The search phrase for the articles you want to retrieve
    CATEGORY: The category of the articles you want to retrieve
    NUMBER_OF_MONTHS: The number of months in the past to search for articles

    The actual version of the code have a test code to run inside Robocorp www.cloud.robocorp.com/
    If you want to use the variables from the setup.py use only this function (Not necessary to copy and paste now, it is already in the code. i'll let you know if I change it):
    ```
    from setup import URL, SEARCH_PHRASE, NUMBER_OF_MONTHS, CATEGORY 
    def main(self) -> None:
        try:
            create_image_folder()
            self.open_website(url=URL)
            self.begin_search(search_phrase=SEARCH_PHRASE)
            self.select_category(categorys=CATEGORY)
            self.sort_newest_news()
            self.set_date_range(NUMBER_OF_MONTHS)
            self.extract_website_data(SEARCH_PHRASE)
        finally:
            self.close_browser()
    ```

Once you have set these variables, you can run the program using Python.


Sample Payload to run inside Robocorp
```
{
  "url": "https://www.nytimes.com/",
  "search_phrase": "business",
  "number_of_months": 1,
  "category": [
    "Arts",
    "Books"
  ]
}
```
