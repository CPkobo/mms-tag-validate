# Validator for Memsource-Bilingual-docx

This tool allows you to validate complex tag/variable strings.

## Usage

### 1. Install pipenv and python-docx

```commandline
pip install pipenv
pipenv install
```

or with global python,

```commandline
pip install python-docx
```

### 2. Put the Memsource-Bilingual-docx into './_local' directory

### 3. Set the RegExp in './reg.txt'

- Example 1
``` txt
\<.*?\>
```

- Example 2
``` txt
\{.*?\}
```

### 4. Execute main.py

```commandline
# with pipenv
pipenv run start
```

```commandline
# with global python
python main.py 
```

Then, you can find the result in the './_local/result.tsv' file. 

This script would take 2-3 minutes per 1 table(1000 rows)