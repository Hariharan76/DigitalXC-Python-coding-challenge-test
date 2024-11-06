import pandas as pd
import re
from collections import Counter

def extract_group_counts(file_path, search_term="Groups"):
    # Load the Excel file
    try:
        excel_data = pd.ExcelFile(file_path)
        df = excel_data.parse(sheet_name="Input Data sheet")
    except ValueError as e:
        print(f"Error: {e}")
        print("Available sheet names:", excel_data.sheet_names)
        return
    
    # Ensure column exists and is accessible
    if 'Additional comments' not in df.columns:
        raise ValueError("The 'Additional comments' column is missing in the file.")
    
    # Define regex pattern for groups (customizable to search_term)
    group_pattern = rf"{search_term}\s*:\s*\[code\]<I>(.*?)<\/I>\[\/code\]"
    
    # Initialize counter
    group_counts = Counter()
    
    # Process each comment for matching groups
    for comment in df['Additional comments'].dropna():
        matches = re.findall(group_pattern, comment)
        for match in matches:
            # Split groups by commas and trim whitespace
            groups = [group.strip() for group in match.split(',')]
            group_counts.update(groups)
    
    # Create output in the specified format
    output_lines = ["Group_name\t\tNumber of occurrences"]
    for group, count in group_counts.items():
        output_lines.append(f"{group}\t\t{count}")
    
    # Write output to a text file
    with open("group_counts_output.txt", "w") as output_file:
        output_file.write("\n".join(output_lines))

    print("Output written to group_counts_output.txt")

# Run the function with the uploaded file path
extract_group_counts('coding challenge test (1).xlsx')
