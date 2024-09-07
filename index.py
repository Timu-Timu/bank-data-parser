import os
import re
from datetime import datetime, timedelta

from bs4 import BeautifulSoup
from group_categories_writer import GroupCategoriesWriter
from results_writer import ResultsWriter

category_writer = GroupCategoriesWriter("data/titles-dictionary.xlsx")
results_writer = ResultsWriter()


def read_html_file(file_path):
    """Read the HTML file and return its content."""
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()


def parse_container(soup):
    """Find the container div and extract relevant data."""
    container_div = soup.find('div', class_=re.compile(r'^operationsstyles__Container'))

    if not container_div:
        print("The specified container div was not found.")
        return

    day_group_divs = container_div.find_all('div')
    for day_group_div in day_group_divs:
        extract_day_info(day_group_div)


def extract_day_info(day_group_div):
    """Extract and print day information and operations."""
    day_info_tag = day_group_div.find('h3')

    if day_info_tag:
        date_string_format = get_date_string(day_info_tag.text.strip())
        print(f"Starting to parse bank data from {date_string_format}")

    operation_divs = day_group_div.find_all('div', class_=re.compile(r'^operationstyles__Row'))
    for operation_div in operation_divs:
        extract_operation_info(operation_div, date_string_format)


def get_date_string(date_str) -> str:
    """
    Get the date string in the format dd.MM.YYYY based on the given month and day.

    Args:
        date_str (str): The day of the month_str and month in Russian (e.g., "Января", "Февраля", etc.).

    Returns:
        str: The date string in the format dd.MM.YYYY.
    """
    now = datetime.now()

    if "Вчера" in date_str:
        date = now - timedelta(days=1)
    else:
        date_str = date_str.replace(' ', '')
        day_str = date_str[:2]
        month_str = date_str[2:]

        month_names = [
            "января", "февраля", "марта", "апреля", "мая", "июня",
            "июля", "августа", "сентября", "октября", "ноября", "декабря"
        ]
        month_index = month_names.index(month_str.lower())
        date = datetime(year=now.year, month=month_index + 1, day=int(day_str))

    return date.strftime("%d.%m.%Y")


def extract_operation_info(operation_div, date_string_format):
    """Extract and print operation information."""
    text_info_div = operation_div.find('div', class_=re.compile(r'^operationstyles__InfoWrapper'))
    amount_div = operation_div.find('div', class_=re.compile(r'^operationstyles__AmountWrapper'))

    if text_info_div and amount_div:
        span_elements = operation_div.find_all('span')

        # Check if there are any span elements
        if span_elements:
            # Iterate through each span element and print its text
            convert_span_info(span_elements, date_string_format)
        else:
            print("No span elements found in this operation div.")
    else:
        print("Required divs not found in this operation div.")


def convert_span_info(span_elements, date_string_format):
    """Check span class and print appropriate information."""
    for span in span_elements:
        if span.get('class'):
            if any("Category" in cls for cls in span.get('class')):
                print(f"Категория: {span.text.strip()}")
                category = span.text.strip()
            elif any("Title" in cls for cls in span.get('class')):
                print(f"Наименование: {span.text.strip()}")
                title = span.text.strip()
            elif any("OperationAmount" in cls for cls in span.get('class')):
                amount = parse_amount(span.text.strip())
                print(f"Цена: {amount:.2f}")  # Print the amount formatted to two decimal places
                results_writer.write_sample_info(title, amount, date_string_format, category_writer)


def parse_amount(amount_str):
    """Parse the amount string into a float."""
    # Remove the currency symbol and whitespace
    amount_str = amount_str.replace('₽', '').strip()

    # Replace space with nothing (remove thousands separator)
    amount_str = amount_str.replace(' ', '')

    # Replace comma with dot (convert decimal separator)
    amount_str = amount_str.replace(',', '.')

    # Convert to float
    try:
        return float(amount_str)
    except ValueError:
        print(f"Error parsing amount: {amount_str}")
        return 0.0  # Return 0.0 or handle the error as needed


if __name__ == "__main__":
    # Set the path to the HTML file
    html_file_path = os.path.join(os.path.dirname(__file__), 'input/data.html')

    # Read and parse the HTML content
    html_content = read_html_file(html_file_path)
    soup = BeautifulSoup(html_content, 'html.parser')

    # Parse the container for data
    parse_container(soup)
    category_writer.write_new_samples_to_excel("data/titles-dictionary.xlsx")
    results_writer.write_results_to_excel()