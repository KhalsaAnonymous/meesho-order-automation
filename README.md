# Meesho Order Automation

This Python script automates the process of fetching order details from Meesho and storing them in an Excel file.

## Introduction

This automation script is designed to streamline the process of retrieving order details from the Meesho platform, a popular online reselling platform, and storing them in an Excel file for further analysis or processing. By automating this task, users can save time and effort, especially when dealing with a large volume of orders.

## Features

- Login to the Meesho supplier panel.
- Retrieve order details using order IDs.
- Store order details (such as sub-order contribution and net order amount) in an Excel file.
- Handle cases where order details are not found gracefully.

## Installation

To run this script, follow these steps:

1. Clone this repository to your local machine:

    ```bash
    git clone https://github.com/KhalsaAnonymous/meesho-order-automation.git
    ```

2. Navigate to the project directory:

    ```bash
    cd meesho-order-automation
    ```

3. Install the required Python packages:

4. Ensure you have Chrome installed on your system, as the script uses Chrome WebDriver for automation.

## Usage

1. Update the `input.xlsx` file with your order IDs.
2. Replace the login credentials and website URLs in the `meesho_order_automation.py` script with your Meesho account details.
3. Run the script using the following command:

    ```bash
    python meesho_order_automation.py
    ```

4. The script will open a Chrome browser, log in to your Meesho account, fetch order details, and store them in the `output.xlsx` file.

## Dependencies

- pandas
- selenium
- webdriver_manager

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, feel free to open an issue or create a pull request.
