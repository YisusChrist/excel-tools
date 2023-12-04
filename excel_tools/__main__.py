#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@file     excel_tools.py
@date     2023-10-13
@version  0.1.0
@license  GNU General Public License v3.0
@url      https://github.com/yisuschrist/excel-tools
@author   Alejandro Gonzalez Momblan (agelrenorenardo@gmail.com)
@desc     Help to work with Excel files in an easy way
"""
import ast
import importlib
import inspect
import locale
import sys

import inquirer
from rich import print
from rich.traceback import install

from .consts import DEBUG, EXIT_FAILURE, EXIT_SUCCESS, LOCALE_FORMAT
from .handlers import print_df, save_excel_file


def get_available_actions():
    """
    Get a list of available actions by inspecting the functions in the script.

    Returns:
        list: A list of available actions.
    """
    # Specify the name of the module ('actions' in this case)
    module_name = "excel_tools.actions"

    # Import the module dynamically
    actions_module = importlib.import_module(module_name)

    # Read the source code of the module
    module_source = inspect.getsource(actions_module)

    # Parse the source code using ast
    module_ast = ast.parse(module_source)

    # Get all functions defined in the module
    return [node.name for node in module_ast.body if isinstance(node, ast.FunctionDef)]


def get_user_action():
    """
    Get the user's choice for the action to perform.

    Returns:
        str: The user's chosen action.
    """
    available_actions = get_available_actions()

    try:
        questions = [
            inquirer.List(
                "action",
                message="What do you wish to do?",
                choices=available_actions,
            ),
        ]
        answers = inquirer.prompt(questions)
        return answers["action"]
    except:
        sys.exit(EXIT_FAILURE)


def handle_user_action(action):
    """
    Handle the user's chosen action.

    Args:
        action (str): The user's chosen action.
    """
    if action not in get_available_actions():
        print("Invalid option.")
        sys.exit(EXIT_FAILURE)

    # Dynamically call the function with the same name as the action
    df = getattr(importlib.import_module("excel_tools.actions"), action)()

    # Print the result
    print_df(df)

    # Ask the user if they want to save the result
    questions = [
        inquirer.Confirm(
            "save",
            message="Do you wish to save the result?",
            default=True,
        ),
    ]
    answers = inquirer.prompt(questions)

    if answers["save"]:
        save_excel_file(df)


def main():
    """Main function."""
    locale.setlocale(locale.LC_ALL, LOCALE_FORMAT)

    # Get the user's chosen action
    action = get_user_action()
    handle_user_action(action)

    sys.exit(EXIT_SUCCESS)


if __name__ == "__main__":
    # Enable rich error formatting in debug mode
    install(show_locals=DEBUG)
    main()
