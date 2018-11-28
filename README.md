# CSVtoXL
Converts a batch of csv files into a single Excel file, with each csv file having a unique sheetname

System: osx

Requirements:
1. Python3
  1a. $ ruby -e "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)" # Installs Homebrew
  1b. $ brew install python # Homebrew installs Python(2)(3) and pip pointing to the Homebrew'd Python3
  1c. $ python --version # Checks if you have python3 installed.
2. openpyxl package
  2a. $ pip3 install openpyxl # installs Python package
  2b. $ python3 # Check if openpyxl was installed in python3.x directory
  2c. >>> import site
  2d. >>> site.getsitepackages('openpyxl')
  
Usage:
1. Place script in any directory
2. Create a 'csvfiles' directory or folder (i.e .../Desktop/csvfiles)
3. Copy your batch csv files into the csvfiles directory or folder
4. Run the script from Terminal
  4a. $ cd Desktop
  4b. $ python3 CSVtoXL.py


