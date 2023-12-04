import sys
import warnings
from tkinter import Tk, filedialog

import inquirer
import pandas as pd
from rich import print as rprint
from rich.console import Console
from rich.markdown import Markdown

from .consts import COLUMN_SUFFIX, DEFAULT_DST_NAME, EXIT_FAILURE


def print_df(df: pd.DataFrame) -> None:
    rprint("Printing DataFrame...")

    # Ask the user if they want to print the DataFrame using rich or pandas
    questions = [
        inquirer.List(
            "print",
            message="How do you wish to print the DataFrame?",
            choices=["rich", "pandas"],
        ),
    ]
    answers = inquirer.prompt(questions)

    if answers["print"] == "rich":
        # Print the DataFrame using rich
        console = Console()

        # Calculate the number of rows to print based on the terminal height
        rows_to_print = console.height - 10
        df = df.head(rows_to_print)

        columns_to_print = 5
        df = df.iloc[:, :columns_to_print]

        md = Markdown(df.to_markdown())
        console.print(md)

        # Print the size of df
        rprint(f"Rows: {df.shape[0]} columns: {df.shape[1]}\n")

    elif answers["print"] == "pandas":
        rprint(df)
    else:
        rprint("Invalid option.")
        sys.exit(EXIT_FAILURE)


def get_excel_file():
    rprint("Select an Excel file")
    # Create a Tk root window
    root = Tk()
    # Hide the main window
    root.withdraw()
    # Ask the user to select a file
    file_path = filedialog.askopenfilename()
    if not file_path:
        rprint("No file selected.")
        sys.exit(EXIT_FAILURE)

    file_name = file_path.split("/")[-1]

    with warnings.catch_warnings(record=True):
        # Read the Excel file
        warnings.simplefilter("always")
        df = pd.read_excel(file_path, engine="openpyxl")

    # Print the size of df
    rprint(f"file '{file_name}' rows: {df.shape[0]} columns: {df.shape[1]}\n")

    return df


def get_excel_files():
    try:
        # Read the Excel files
        return get_excel_file(), get_excel_file()
    except FileNotFoundError as e:
        rprint(e)
        sys.exit(EXIT_FAILURE)


def get_key(df1: pd.DataFrame, df2: pd.DataFrame) -> str:
    """
    Get the key column to use to check the registers.

    Args:
        df1 (pd.DataFrame): The first DataFrame.
        df2 (pd.DataFrame): The second DataFrame.

    Returns:
        str: The key column.
    """
    # TODO: Add support for multiple keys
    # TODO: Add support for multiple arguments to extract the key from multiple dataframes
    # Get the list of common columns between df1 and df2
    common_columns = df1.columns.intersection(df2.columns).tolist()

    # Ask the user to select a column to use as key
    questions = [
        inquirer.List(
            "key",
            message="What column do you wish to use to check the registers?",
            choices=common_columns,
        ),
    ]
    answers = inquirer.prompt(questions)
    return answers["key"]


def ask_file_name():
    rprint(
        "Enter the name of the desired output "
        f"file (default is '{DEFAULT_DST_NAME}'): ",
        end="",
    )
    dst_name = input().strip() or DEFAULT_DST_NAME

    if not dst_name.endswith(".xlsx"):
        dst_name += ".xlsx"

    return dst_name


def save_excel_file(df: pd.DataFrame) -> None:
    """
    Save the result to a new Excel file.

    Args:
        df (pd.DataFrame): The DataFrame to save.
    """
    max_retries = 3
    retry_count = 0
    saved = False

    try:
        dst_name = ask_file_name()
    except KeyboardInterrupt:
        rprint("Operation cancelled.")
        sys.exit(EXIT_FAILURE)

    while not saved and retry_count < max_retries:
        try:
            # Save the result to a new Excel file
            df.to_excel(dst_name, index=False)
            saved = True
            rprint(f"Result saved as '{dst_name}'")
        except PermissionError as e:
            rprint(e)
            retry_count += 1
            if retry_count < max_retries:
                input(
                    "Please make sure the Excel file is closed and press Enter to retry."
                )
            else:
                rprint("Maximum number of retries reached. Unable to save the result.")


def verify_values():
    """
    Read the values from the columns with name X and X (Num). The column X
    contains string values and the column X (Num) contains numeric values. The
    function must check if the values in the column X (Num) are the same as the
    values in the column X but with the last character removed. If the values
    are not the same, the function must return the rows that do not match.

    Note that there may be multiple pair of columns that match this condition
    and the function must do the verification for all of them.

    For example, we can have a column called 'Purchasable Item' with a row
    containing the value 564 (which is a string) and another column called
    'Purchasable Item (Num)' with a row containing the value 564,00 (which is
    a number). In this case, those 2 values are correct as the conversion from
    string to number is correct. However, in this other example we can have
    the first column with value 1,32 and the second column with value 132,00
    which are not correct as the conversion from string to number is not
    correct. That row must be returned by the function.

    Returns:
        pd.DataFrame: The rows that do not match.
    """
    # Read the Excel file
    df = get_excel_file()
    # Get the key column

    column_suffix_len = len(COLUMN_SUFFIX)

    columns_to_check = []
    for column_name in df.columns:
        if column_name.endswith(COLUMN_SUFFIX):
            # Extract the corresponding "X" column name
            x_column_name = column_name[:-column_suffix_len]

            # Check if the "X" column exists in the DataFrame
            if x_column_name in df.columns:
                columns_to_check.append(x_column_name)

    rprint("The following columns will be checked:")
    rprint(columns_to_check)

    mismatched_rows = []

    # Filter the DataFrame to only contain the row where the value of the
    # column Short Code is ES1S60WRMBZ
    df = df[df["Short Code"] == "ES1S60WRMBZ"]

    # Iterate through the columns of the DataFrame to find matching pairs
    for column_name in columns_to_check:
        # Check for matching pairs of columns
        x_column_name = column_name + COLUMN_SUFFIX

        rprint(f"Checking columns '{column_name}' and '{x_column_name}'...")

        # Get the values of the columns and compare the type
        str_column_1 = df[column_name].astype(str).str.replace(",", "")
        rprint("str_column_1:")
        rprint(str_column_1)

        # Convert the values of the column column_name to numeric
        num_column = pd.to_numeric(str_column_1, errors="coerce")
        rprint(f"num_column:")
        rprint(num_column)

        str_column_2 = df[x_column_name]
        rprint("str_column_2:")
        rprint(str_column_2)

        # Check if the values of the columns are the same
        mismatched_rows += df[(num_column != str_column_2)].to_dict(orient="records")

    if not mismatched_rows:
        rprint("No mismatches found.")
    else:
        # Remove duplicates
        mismatched_rows = pd.DataFrame(mismatched_rows).drop_duplicates()
        rprint(f"Found {len(mismatched_rows)} rows with mismatches.")

        try:
            input("Press Enter to continue...")
        except KeyboardInterrupt:
            rprint("Operation cancelled.")
            sys.exit(EXIT_FAILURE)

    return pd.DataFrame(mismatched_rows)
