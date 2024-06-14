# Import the os library for accessing operating system functions and file system operations
import os

# Import the fitz library from the PyMuPDF package for processing PDF documents
import fitz

# Import the pandas library for data manipulation and exporting to Excel files
import pandas as pd

def count_word_occurrences(folder_path, search_words):
    # Initialize an empty word count dictionary for each search word
    word_counts = {word: [] for word in search_words}
    # Initialize an empty list for the names of found papers
    paper_names = []

    # Iterate over all files in the specified folder
    for filename in os.listdir(folder_path):
        # Check if the file is a PDF file
        if filename.endswith('.pdf'):
            # Construct the full file path
            file_path = os.path.join(folder_path, filename)
            # Open the PDF document using the fitz library
            doc = fitz.open(file_path)
            # Initialize an empty set for each PDF document to count each word only once per document
            found_words = set()
            # Iterate over each search word
            for word in search_words:
                # Initialize a variable to determine if the word was found in the current document
                word_found = False
                # Iterate over each page of the PDF document
                for page in doc:
                    # Extract the text from the page
                    text = page.get_text()
                    # Check if the word appears in the text and it has not been found previously in the current document
                    if word.lower() in text.lower() and word not in found_words:
                        # Increase the word count for the current word
                        word_counts[word].append(1)
                        # Add the word to the set of found words
                        found_words.add(word)
                        # Set the variable to True to indicate that the word was found
                        word_found = True
                        break
                # If the word was not found, append a 0 to the word count
                if not word_found:
                    word_counts[word].append(0)
            # Add the filename to the list of paper names
            paper_names.append(filename)

    # Return the word count dictionary and the list of paper names
    return word_counts, paper_names

# Specified folder path where the PDF documents are stored
folder_path = r'C:\Users\neues\OneDrive - FH Muenster\Dokumente\FH Münster\FH Münster_SS23\Masterarbeit\Konzeptmatrix und Literatur\Literaturexporte\Kategorisierte Paper'
# Search words to look for
search_words = ['Deep Learning', 'Machine Learning', 'Artificial Intelligence', 'Crawl', 'Crawling', 'Scrape', 'Scraping', 'Web Application', 'User Interface']

# Call the function to count word occurrences and obtain the paper names
word_counts, paper_names = count_word_occurrences(folder_path, search_words)

# Create a data dictionary with paper names and word counts
data = {'Paper': paper_names}
for word, counts in word_counts.items():
    data[word] = counts

# Create a DataFrame from the data dictionary
df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
output_file = 'word_counts.xlsx'
df.to_excel(output_file, index=False)

# Print a message indicating that the word count has been saved in the Excel file
print(f"Word counts have been saved in the Excel file '{output_file}'.")
