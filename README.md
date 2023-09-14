# Stock Data Analysis Tool

This is a VBA script that analyzes one year of stock data in an Excel workbook. It retrieves, processes, and calculates various stock-related values and formats the data according to specified requirements.

## Table of Contents

- [Requirements](#requirements)
- [How to Use](#how-to-use)
- [Features](#features)
- [License](#license)

## Requirements

The VBA script meets the following requirements:

### Retrieval of Data

- The script loops through one year of stock data and reads/stores the following values from each row:
  - Ticker symbol
  - Volume of stock
  - Open price
  - Close price

### Column Creation

- On the same worksheet as the raw data, or on a new worksheet, all columns were correctly created for:
  - Ticker symbol
  - Total stock volume
  - Yearly change ($)
  - Percent change

### Conditional Formatting

- Conditional formatting is applied correctly and appropriately to the yearly change column.
- Conditional formatting is applied correctly and appropriately to the percent change column.

### Calculated Values

- All three of the following values are calculated correctly and displayed in the output:
  - Greatest % Increase
  - Greatest % Decrease
  - Greatest Total Volume

### Looping Across Worksheet

- The VBA script can run on all sheets successfully.

## How to Use

1. Open your Excel workbook containing the stock data you want to analyze.

2. Press `ALT` + `F11` to open the Visual Basic for Applications (VBA) editor.

3. In the VBA editor, insert a new module (if not already present).

4. Copy and paste the entire VBA script into the module.

5. Close the VBA editor and return to your Excel workbook.

6. Make sure your stock data is in a format similar to the one specified in the requirements.

7. Press `ALT` + `F8` to open the "Macro" dialog box.

8. Select the script you just added and click "Run."

9. The script will process the data, create the necessary columns, apply conditional formatting, and display the calculated values.

## Features

- **Data Retrieval**: The script efficiently loops through the stock data to retrieve essential values.

- **Column Creation**: It creates new columns as required, including ticker symbols, total stock volume, yearly change ($), and percent change.

- **Conditional Formatting**: Proper conditional formatting is applied to highlight positive and negative changes in the yearly and percent change columns.

- **Calculated Values**: The script accurately calculates and displays the greatest % increase, greatest % decrease, and greatest total volume for the given dataset.

- **Looping Across Worksheets**: The script is designed to work seamlessly across all sheets in the workbook, ensuring consistency in data analysis.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
