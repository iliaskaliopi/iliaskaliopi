import pandas as pd
import random
import string
import os


def generate_random_filename(extension=".html", length=8):
    """Generate a random unique alpharithmetic filename with a given extension."""
    chars = string.ascii_letters + string.digits
    return ''.join(random.choice(chars) for _ in range(length)) + extension


def process_csv(input_csv, template_file, output_csv):
    # Read the CSV file into a DataFrame
    df = pd.read_excel(input_csv)

    # Ensure there is a 'text' column in the DataFrame
    if 'Name' not in df.columns:
        raise ValueError("CSV must contain a 'text' column.")

    # Read the template file
    with open(template_file, 'r', encoding='utf-8') as file:
        template_content = file.read()

    # Create a list to store the new filenames
    filenames = []

    # Process each row in the DataFrame
    for index, row in df.iterrows():
        # Replace the placeholder with the actual text
        new_content = template_content.replace("[KALESMENOS]", row['Name'])

        # Generate a random unique filename
        new_filename = generate_random_filename()

        # Write the new content to the new file
        with open(new_filename, 'w', encoding='utf-8') as file:
            file.write(new_content)

        # Append the filename to the list
        filenames.append(new_filename)

    # Add the new filenames as a column in the DataFrame
    df['link'] = [f"C:\\Users\\apost\\Documents\\weding\\{f}" for f in filenames]

    # Save the modified DataFrame to a new CSV file
    df.to_excel(output_csv, index=False)


# Example usage
input_csv = 'lista.xlsx'  # Path to the input CSV file
template_file = 'index.html'  # Path to the template HTML file
output_csv = 'final.xlsx'  # Path to the output CSV file

process_csv(input_csv, template_file, output_csv)