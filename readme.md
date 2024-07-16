### 1. Install Python:

- macOS usually comes with Python pre-installed, but it's better to install a newer version.
- Go to the official Python website: https://www.python.org/downloads/
- Click on the "Download Python" button. It should automatically suggest the latest version for macOS.
- Once downloaded, open the installer package (.pkg file).
Follow the installation wizard, accepting the default options.


### 2. Verify Python installation:

- Open Terminal (you can find it in Applications > Utilities > Terminal)
- Type python3 --version and press Enter
- You should see the Python version number. If you do, Python is installed correctly.


### 3. Install pip (Python package manager):

- In the Terminal, type the following command and press Enter:
```curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py```

- Once the download is complete, run:
```python3 get-pip.py```



### 4. Install required libraries:

- In the Terminal, run:
```pip3 install openpyxl ebooklib```

### 5. Run the script:

- Download the script from the repository.
- Open the Terminal and navigate to the folder where the script is saved.
- Save your .xlsx file in the same folder.
- Run the script by typing:
```python3 xlsx_to_epub.py```


## about the xlsx format:
- The script expects the .xlsx file to have the following structure:
  - The first row should contain the column names.
  - The first column should contain the chapter titles.
  - The second column should contain the text content.


   for example:
  | Chapter | Content |
  | --- | --- |
  | Chapter 1 | This is the content of chapter 1. |
  | Chapter 2 | This is the content of chapter 2. |