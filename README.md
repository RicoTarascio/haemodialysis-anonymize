# haemodialysis-anonymize

Python service to anonymize files coming from haemodialysis

## Dependencies

Requires openpyxl >= 3.1.2

## Usage

It is advisable to install this in a Python virtualenv without system packages.

To test full functionalities run

`import anonymize`

`anonymize.anonymize()`

The function will:

1. Create the folder tests/anonymized
2. Run through all the files in tests/original (if the folder does not exist one will be created and the process is going to be interrupted)
3. For each of them generate a unique ID based on the name of the file (naming convention is name-surname-DD_MM_YYYY ). A copy of the file with ID-currentyear as name will be added in tests/anonymized
4. Create anon_entries.xlsx in tests/anonymized
5. For each ID generated create an entry in anon_entries.xlsx as the following: name, surname, birth_date, ID
