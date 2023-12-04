[![License: GPL v3.0](https://img.shields.io/badge/License-GPL%20v3.0-blue.svg)](https://opensource.org/licenses/GPL-3.0)
[![Python Version](https://img.shields.io/badge/Python-3.6%2B-blue)](https://www.python.org/downloads/)
[![Pandas Version](https://img.shields.io/badge/pandas-1.3.3%2B-brightgreen)](https://pandas.pydata.org/)
[![Inquirer Version](https://img.shields.io/badge/inquirer-2.7.0%2B-brightgreen)](https://github.com/magmax/python-inquirer)
[![Rich Version](https://img.shields.io/badge/rich-10.12.0%2B-brightgreen)](https://github.com/willmcgugan/rich)

Excel Tools is a Python script designed to simplify working with Excel files in an easy way. It provides functionalities such as removing duplicated rows, comparing two Excel files, verifying values, and merging two Excel files.

Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Usage](#usage)
- [License](#license)
- [Acknowledgements](#acknowledgements)
- [Contributing](#contributing)

# Features

- **Remove Duplicated Rows**: Easily identify and remove duplicated rows from an Excel file.
- **Compare Two Excel Files**: Identify rows present in the first Excel file but not in the second.
- **Verify Values**: Check and verify values in specified columns for correctness.
- **Merge Two Excel Files**: Combine data from two Excel files based on common columns.

# Requirements

- Python 3.6 or later
- Pandas 1.3.3 or later
- Inquirer 2.7.0 or later
- Rich 10.12.0 or later

# Usage

1. Clone the repository:

   ```bash
   git clone https://github.com/yisuschrist/excel-tools
   cd excel-tools
   ```

2. Install the required dependencies:

   ```bash
   poetry install --no-dev
   ```

3. Run the script:

   ```bash
   poetry run python -m excel_tools
   ```

4. Follow the prompts to perform desired actions.

# License

This project is licensed under the [GNU General Public License v3.0](https://opensource.org/licenses/GPL-3.0).

# Acknowledgements

This project utilizes the following open-source projects:

- [Pandas](https://pandas.pydata.org)
- [Inquirer](https://github.com/magmax/python-inquirer)
- [Rich](https://github.com/Textualize/rich)
- [Poetry](https://python-poetry.org)

# Contributing

Contributions are welcome! Feel free to open issues or pull requests.
