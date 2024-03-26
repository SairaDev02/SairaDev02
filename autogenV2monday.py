"""
autogenV2monday.py

Author: sairadev02

This script generates Word documents for the current Monday and the next 4 days. 
It creates the documents in the same directory as the script and logs the process.

The script can only be executed on Mondays. If it is executed on any other day, it will raise a ValueError.

The file names for the Word documents are in the format: 77988_dailyreport_YYYYMMDD.docx

Modify the format for individual needs at line 60. (IMPORTANT)
"""

# Import necessary modules

import os # The os module provides a way to interact with the operating system
import datetime # The datetime module provides classes for manipulating dates and times
import win32com.client # The win32com.client module provides access to COM objects in Windows
import pythoncom # The pythoncom module provides support for COM (Component Object Model) in Python
import logging # The logging module provides a way to log messages for debugging and monitoring
from concurrent.futures import ThreadPoolExecutor # The concurrent.futures module provides a high-level interface for asynchronously executing functions

# Set the directory path to the current location of the file
# This will ensure that the Word documents are saved in the same directory as the script
directory = os.path.dirname(os.path.abspath(__file__))

# Create the directory if it doesn't exist
# This will ensure that the Word documents are saved in the correct location
if not os.path.exists(directory):
    os.makedirs(directory)

# Get the current date
# This will be used to determine the file names for the Word documents
today = datetime.date.today()

# Check if the script is executed on a Monday
if today.weekday() != 0:  # Monday is 0
    raise ValueError("This script can only be executed on Mondays.")

# Generate file names and paths for the current and next 4 days
# The file names will be in the format: 77988_dailyreport_YYYYMMDD.docx
# The file paths will be in the format: directory/77988_dailyreport_YYYYMMDD.docx
file_names = []
file_paths = []

# Loop through the next 5 days (including today)
# This will generate file names and paths for the current Monday and the next 4 days
for i in range(5):

    # Calculate the date for the current iteration
    # This will be the current date plus the iteration offset
    # For example, if i = 1, it will be the next day
    # If i = 0, it will be the current day
    # If i = -1, it will be the previous day
    date = today + datetime.timedelta(days=i)

    # Generate the file name based on the date (Modify the format as needed)
    file_name = f"77988_dailyreport_{date.strftime('%Y%m%d')}.docx"

    # Generate the file path based on the directory and file name
    file_path = os.path.join(directory, file_name)

    # Check if the file already exists
    if os.path.exists(file_path):
        
        # If the file already exists, raise an error
        # This will prevent overwriting existing files
        raise FileExistsError(f"File already exists: {file_path}")
    
    # Append the file name and path to the lists
    # This will be used to create the Word documents later
    file_names.append(file_name)
    file_paths.append(file_path)

# Set up logging
# This will create a log file in the same directory as the script
logging.basicConfig(filename=os.path.join(directory, "autogenV2-monday.log"), level=logging.INFO)
logging.info(f"Creating Word documents for the following days: {file_names}")
logging.info(f"Files will be saved to the directory: {directory}")
logging.info("Starting document creation...")

# Set up text feedback in terminal
# This will provide real-time feedback on the progress of the script
print(f"Creating Word documents for the following days: {file_names}")
print(f"Files will be saved to the directory: {directory}")
print("Starting document creation...")

# Function to create a Word document at the specified file path
# This function will be executed in parallel for each file path
# It uses the win32com.client module to interact with Microsoft Word
# The pythoncom module is used to initialize and uninitialize the COM environment
def create_word_document(file_path):
    """
    Creates a Word document at the specified file path.

    Parameters:
    file_path (str): The path where the Word document should be created.

    Returns:
    bool: True if the document was created successfully, False otherwise.
    """
    try:
        # Initialize the COM environment
        # This is required for interacting with Microsoft Word
        pythoncom.CoInitialize()

        # Create a new Microsoft Word instance
        # This will open a new Word application window
        word = win32com.client.Dispatch("Word.Application")
        word.DisplayAlerts = 0  # Disable alerts

        # Create a new document
        # This will open a new document window in the Word application
        doc = word.Documents.Add()

        # Save the document with the specified file name and path
        # The document will be saved in the specified directory
        doc.SaveAs(file_path)

        # Close the document and Word application
        # The changes will be saved because we already saved the document
        doc.Close()
        word.Quit()

        # Uninitialize the COM environment
        # This is required to release the resources
        pythoncom.CoUninitialize()

        # Log and print the success message
        logging.info(f"Document created successfully: {file_path}")
        print(f"Document created successfully: {file_path}")

    # Handle any exceptions that occur during the document creation
    except Exception as e:
        # Log and print the error message
        logging.error(f"Error creating document: {file_path}")
        print(e)

        # Return False to indicate that the document creation failed
        return False
    
    # Return True to indicate that the document creation was successful
    return True

# Use ThreadPoolExecutor to create the Word documents in parallel
# This will speed up the process significantly
with ThreadPoolExecutor() as executor:

    # Map the function to create the Word documents with the file paths
    # This will execute the function in parallel for each file path
    # The results will indicate whether each document was created successfully
    results = executor.map(create_word_document, file_paths)

# Check if all documents were created successfully
if all(results):

    # Log and print completion messages
    # This will indicate that the script has finished running
    logging.info("Document creation completed.")
    logging.info("-" * 120) # Add a separator for better readability
    print("Document creation completed.")
else:
    logging.error("Some documents could not be created.")
    print("Some documents could not be created.")

print("Have a great day!") # End of script message