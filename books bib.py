import requests
import pandas as pd

def fetch_book_info(isbn):
    
    url = f'https://www.googleapis.com/books/v1/volumes?q=isbn:{isbn}'

    response = requests.get(url)
    data = response.json()

    if 'items' in data:
        book_info = data['items'][0]['volumeInfo']
        return book_info
    else:
        return None

def save_to_excel(book_data, output_file='book_data.xlsx'):
    df = pd.DataFrame(book_data)
    df.to_excel(output_file, index=False)
    print(f'Data saved to {output_file}')

def main():
    # Replace 'your_input_file.xlsx' with the name of your Excel file containing ISBNs
    input_file = 'Book_bib.xlsx'
    output_file = 'book_data.xlsx'

    # Read ISBNs from the Excel file
    isbn_data = pd.read_excel(input_file, header=None, names=['ISBN'])
    isbns = isbn_data['ISBN'].tolist()

    book_data = []

    for isbn in isbns:
        print(f'Fetching data for ISBN: {isbn}')
        book_info = fetch_book_info(isbn)
        if book_info:
            book_data.append(book_info)
        else:
            print(f'No data found for ISBN: {isbn}')

    save_to_excel(book_data, output_file)

if __name__ == "__main__":
    main()
