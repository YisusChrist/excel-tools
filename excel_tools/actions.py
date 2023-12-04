import sys

import inquirer
import pandas as pd
from rich import print

from .consts import EXIT_FAILURE
from .handlers import get_excel_file, get_excel_files, get_key, print_df


def remove_duplicates() -> pd.DataFrame:
    """
    Remove duplicated rows.

    Returns:
        pd.DataFrame: The DataFrame without duplicated rows.
    """
    # Read the Excel file
    df = get_excel_file()
    # Get the key column
    key = "Número"
    # Check if there are duplicated rows with the same key
    duplicated = df[df.duplicated(subset=[key], keep=False)]

    print(f"Number of duplicated rows: {duplicated.shape[0]}")
    if duplicated.shape[0] > 0:
        print("Duplicated rows:")
        print_df(duplicated)

    return df.drop_duplicates(subset=[key], keep="first")


def compare_files() -> pd.DataFrame:
    """
    Compare two Excel files.

    Returns:
        pd.DataFrame: The rows of the first Excel file that are not present in the second Excel file.
    """
    # Read the Excel files
    df1, df2 = get_excel_files()
    # Get the key column
    key = get_key(df1, df2)
    # Find rows with key not present in df2
    df1_not_in_df2 = df1[~df1[key].isin(df2[key])]

    return df1_not_in_df2


def get_matches(df1_unique_columns: set, df2_unique_columns: set) -> pd.DataFrame:
    """
    Get the most similar matches.

    Args:
        df1_unique_columns (set): The columns of the first Excel file that are not present in the second Excel file.
        df2_unique_columns (set): The columns of the second Excel file that are not present in the first Excel file.

    Returns:
        pd.DataFrame: The most similar matches.
    """
    # Get the most similar matches
    matches = pd.DataFrame(
        {
            "df1_column": list(df1_unique_columns),
            "df2_column": list(df2_unique_columns),
            "similarity": 0,
        }
    )

    for i, row in matches.iterrows():
        similarity = 0
        for char1, char2 in zip(row["df1_column"], row["df2_column"]):
            if char1 == char2:
                similarity += 1
            else:
                break
        matches.at[i, "similarity"] = similarity

    return matches.sort_values("similarity", ascending=False)


def merge_excel_files() -> pd.DataFrame:
    """
    Merge two Excel files based on specified columns.

    Reads the first Excel file ("result.xlsx") and asks the user to select the
    second Excel file. Compares the specified columns ("Finales", "Código",
    "Tipo") and retrieves the matching rows.

    Returns:
        pd.DataFrame: DataFrame containing the matched rows.
    """
    try:
        # Read the Excel files
        df1, df2 = get_excel_files()

        # Calculate the intersection of the columns of both DataFrames
        common_columns = df1.columns.intersection(df2.columns).tolist()

        # Prompt the user to select the columns to compare
        inquirer_columns = [
            inquirer.Checkbox(
                "columns",
                message="Select the columns to compare",
                choices=common_columns,
            ),
        ]
        answers = inquirer.prompt(inquirer_columns)
        columns = answers["columns"]

        # Compare the columns and retrieve the matching rows
        matched_rows = pd.merge(df1, df2, on=columns)

        # Return the DataFrame containing matched rows
        return matched_rows
    except Exception as e:
        print(e)
        sys.exit(EXIT_FAILURE)
