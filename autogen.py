import os
import datetime
import win32com.client

# Set the directory path to the current location of the file
directory = os.path.dirname(os.path.abspath(__file__))

# Create the directory if it doesn't exist
if not os.path.exists(directory):
    os.makedirs(directory)

# Get the current date
today = datetime.date.today()

# Format the file name with the specified naming convention
file_name = f"77988_dailyreport_{today.strftime('%Y%m%d')}.docx"

# Join the directory path and file name
file_path = os.path.join(directory, file_name)

# Create a new Microsoft Word instance
word = win32com.client.Dispatch("Word.Application")
word.DisplayAlerts = 0  # Disable alerts

# Create a new document
doc = word.Documents.Add()

# Save the document with the specified file name and path
doc.SaveAs(file_path)

# Close the document and Word application
doc.Close()
word.Quit()