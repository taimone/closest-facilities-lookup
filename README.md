Closest Facilities Lookup Script
This Python script processes employee and facility data to generate an output Excel file that provides information about the closest facilities to each employee's location.

Prerequisites
Python: This script requires Python 3.9 or higher.

Python Libraries: The following Python libraries are required:

asyncio
aiohttp
pandas
openpyxl
You can install these libraries using pip:

pip install asyncio aiohttp pandas openpyxl
Google Maps Distance Matrix API Key: The script uses the Google Maps Distance Matrix API to fetch the distances between the employee zip codes and the facility zip codes. You'll need to obtain an API key from the Google Cloud Console and replace the API_KEY variable in the script with your own API key.

Input Files
The script expects two input CSV files:

input.csv: This file should contain the employee information, including the employee name and zip code. The file should have a column named "Name" and a column named "Employee Zip".
facilities.csv: This file should contain the facility information, including the facility zip code and airport code. The file should have a column named "Facility Zip" and a column named "Airport Code".
Make sure to place these CSV files in the same directory as the Python script.

Output File
The script will generate an output Excel file named output.xlsx in the same directory as the script.

Script Execution
To run the script, execute the following command in a terminal or command prompt:

python script_name.py
Replace script_name.py with the actual name of the Python file containing the script.

Script Functionality
The script performs the following tasks:

Network and API Connectivity Tests:

It first tests the network connectivity by checking if it can connect to Google.com.
It then tests the connection to the Google Maps Distance Matrix API to ensure the API key is valid.
Data Processing:

The script reads the input.csv and facilities.csv files.
It processes the distances between each employee's zip code and the facility zip codes using the Google Maps Distance Matrix API.
The script batches the facility zip codes to optimize the API requests and handles any cases where distance data is not available.
The processed distances are sorted to identify the 3 closest facilities for each employee.
Output Generation:

The script generates an output Excel file named output.xlsx in the same directory as the script.
The output file contains the following information for each employee:
Employee Name
Employee Zip Code
The 3 closest facility zip codes
The distances (in miles) to those 3 closest facilities
The airport code for each of the 3 closest facilities
Expected Returns
The script generates an output Excel file named output.xlsx that contains the information about the 3 closest facilities for each employee, including the facility zip code, distance in miles, and airport code.
