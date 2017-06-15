Python module designed to add TEMCAGT section notes to a google spreadsheet

Requires:

    - gspread
    - oauth2client

Two environment variables are needed:

    - TEMCA_GOOGLE_SPREADSHEET_KEY:
        - This is the spreadsheet specific key that is found in the
          url of the spreadsheet
    - GOOGLE_APPLICATION_CREDENTIALS
        - This is the path to the credentials json that is required
          by gspread
        - How to obtain credentials:
            - http://gspread.readthedocs.io/en/latest/oauth2.html

Arguments:
    
    - source:
        - the location of files to parse and add to the spreadsheet. [REQUIRED]
    - sheet_name:
        - the name of the worksheet to add to (not the spreadsheet name). [REQUIRED]
    - initial:
        - this flag will have the program run the initial update to
          an empty spreadsheet. This must only be run with and EMPTY
          spreadsheet, and will populate the spreadsheet with every
          acceptable file found in the source.
    - update:
        - this flag will have the program run the update function
          that will identify the slot number of the last populated row
          in the spreadsheet. It will find all the files within the source
          dir that have slot numbers greater than the last slot found,
          and populate the spreadsheet with those notes.
    - stop_number:
        - this variable is used to denote a slot number to restrict
          the addition of any notes at or above this slot number. [OPTIONAL]
    - note_keyword:
        - this variable is to specify a keyword that is found only
          within the file names that are to be uploaded into the
          spreadsheet. [REQUIRED]

Notes:

    - When running the update function, if there is a new section note that has the same slot number as the 
      last populated cell in the spreadsheet, it will be skipped and needs to be manually added into the
      spreadsheet
    - Example keywords would be: 'r47', 'ldms2, etc. Any file that includes this keyword will be parsed
    - Gspread currently has an issue with the certifi module, and requires that certifi version 
      2015.4.28 be used in order to work properly
