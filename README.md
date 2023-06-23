# Diesel Fuel Prices Scraper

This project scrapes weekly average diesel fuel prices in Vancouver, BC, from an Excel file available online, and renders it as an HTML table with a dark theme.

## Features

- Retrieves data from an Excel file hosted online.
- Extracts relevant data (Week Ending and Price) from the Excel file.
- Renders the data in an HTML table.
- The HTML table is styled with a dark theme and centered on the page.
- The hyperlink to the source Excel file is styled to blend with the dark theme.

## Requirements

- Node.js
- npm

## Dependencies

- axios
- exceljs
- stream

## Setup and Installation

1. Clone this repository to your local machine.

    ```sh
    git clone <repository-url>
    ```

2. Navigate to the project directory.

    ```sh
    cd <project-directory>
    ```

3. Install the required dependencies.

    ```sh
    npm install
    ```

4. Run the script.

    ```sh
    node index.js
    ```

5. The script will fetch the data and render it as an HTML table. You should see the output in the console.

## Code Structure

- `index.js`: This file contains the main script for fetching and processing the Excel data, and rendering the HTML output.

## Customization

If you wish to customize the appearance of the HTML output, you can modify the CSS styles in the `index.js` file.

## License

This project is licensed under the MIT License.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## Acknowledgements

- Excel file source: [Kalibrate](https://charting.kalibrate.com/WPPS/Diesel/Retail%20(Incl.%20Tax)/WEEKLY/2023/Diesel_Retail%20(Incl.%20Tax)_WEEKLY_2023.xlsx)

## Contact

If you want to contact the author, you can reach out at `<your-email-address>`.
