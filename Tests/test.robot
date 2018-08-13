*** Settings ***
Library           ../ExcellentLibrary.py
Library           OperatingSystem
Library           Collections
# Test Setup        Setup test
# Test Teardown     Teardown test



*** Variables ***
${PROPER EXCEL FILE}                ${CURDIR}${/}Proper Excel File.xlsx
${WEIRD EXCEL FILE}                 ${CURDIR}${/}Weird Excel File.xlsx
${WORKBOOK DUMMY FILENAME}          New Book.xslx



*** Test Cases ***
Opening workbook without alias (defaulting to filepath)
    Open workbook  ${PROPER EXCEL FILE}
    Switch workbook  ${PROPER EXCEL FILE}
    Close workbook  ${PROPER EXCEL FILE}

Opening workbook with alias
    Open workbook  ${PROPER EXCEL FILE}  alias=first excel file
    Switch workbook  first excel file
    Run keyword and expect error  UnknownWorkbookException*
    ...  Switch workbook  ${PROPER EXCEL FILE}
    Close workbook  first excel file

Get row and column count
    Open workbook  ${PROPER EXCEL FILE}  alias=first excel file
    Switch sheet  Sheet 1 (with header)
    ${row count}=  Get row count
    ${column count}=  Get column count
    Should be equal  ${row count}  ${4}
    Should be equal  ${column count}  ${3}
    Switch sheet  Sheet 2 (no header)
    ${row count}=  Get row count
    ${column count}=  Get column count
    Should be equal as integers  ${row count}  ${2}
    Should be equal as integers  ${column count}  ${3}
    Close workbook  first excel file

Reading cell values
    Open workbook  ${PROPER EXCEL FILE}  alias=first excel file
    Switch sheet  Sheet 1 (with header)  # Important!
    ${a4}=  Read from cell  a1:A4
    ${b4}=  Read from cell  (2, 4)
    Should be equal  ${a4}  ${SPACE}${SPACE}First name with leading spaces
    Should be equal  ${b4}  ${SPACE}Last name with whitespace surrounding it${SPACE}${SPACE}${SPACE}
    ${a4}=  Read from cell  1,4  trim=${TRUE}
    ${b4}=  Read from cell  a1:B4  trim=${TRUE}
    Should be equal  ${a4}  First name with leading spaces
    Should be equal  ${b4}  Last name with whitespace surrounding it
    Close workbook  alias=first excel file

Reading cells while switching sheets
    Open workbook  ${PROPER EXCEL FILE}  alias=first excel file
    Switch sheet  Sheet 1 (with header)  # Important!
    ${a1}=  Read from cell  A1
    Should be equal  ${a1}  First name
    Switch sheet  Sheet 2 (no header)
    ${a1}=  Read from cell  A1
    Should be equal  ${a1}  Michiel
    Close workbook  alias=first excel file

Opening several workbooks and switching between them
    Open workbook  ${PROPER EXCEL FILE}  alias=first excel file
    Open workbook  ${WEIRD EXCEL FILE}  alias=second excel file
    Switch workbook  first excel file
    Switch sheet  Sheet 2 (no header)
    ${a1}=  Read from cell  A1
    Should be equal  ${a1}  Michiel
    Switch workbook  second excel file
    Switch sheet  Data (empty header)
    ${a1}=  Read from cell  A1
    Should be equal  ${a1}  ${NONE}
    Close workbook  alias=first excel file
    Close workbook  alias=second excel file

Opening workbook makes it active
    Open workbook  ${PROPER EXCEL FILE}  alias=first excel file
    Open workbook  ${WEIRD EXCEL FILE}  alias=second excel file
    Switch sheet  Data (empty header)  # Present only in the second file.
    Close workbook  alias=first excel file
    Close workbook  alias=second excel file

Closing active workbook
    Open workbook  ${PROPER EXCEL FILE}  first excel file
    Open workbook  ${WEIRD EXCEL FILE}  second excel file
    Switch sheet  Data (empty header)  # Present only in the second file.
    Close workbook  # Implicitly activates next workbook
    Switch sheet  Sheet 1 (with header)  # Present only in the second file.
    Close workbook  # No alias supplied: default to current.

Creating a new workbook and check if it doesn't overwrite existing files by default
    ${workbook exists}=  Run keyword and return status
    ...  File should exist  ${WORKBOOK DUMMY FILENAME}
    Run keyword if  ${workbook exists}
    ...  Remove file  ${WORKBOOK DUMMY FILENAME}
    Create workbook  ${WORKBOOK DUMMY FILENAME}

    Run keyword and expect error  FileAlreadyExists*
    ...  Create workbook  ${WORKBOOK DUMMY FILENAME}
    Close workbook

Creating a new workbook and overwrite existing file
    Create workbook  ${WORKBOOK DUMMY FILENAME}  overwrite_file_if_exists=${TRUE}  alias=book
    Switch workbook  book

Create and remove new sheet
    Open workbook  ${PROPER EXCEL FILE}  first excel file
    Run keyword and ignore error
    ...  Remove sheet  TEST
    Run keyword and expect error  SheetNotFoundException*
    ...  Remove sheet  TEST
    Create sheet  TEST
    Run keyword and expect error  SheetExistsAlreadyException*
    ...  Create sheet  TEST
    Save
    Close workbook
    Open workbook  ${PROPER EXCEL FILE}
    Remove sheet  TEST
    Save
    Close workbook

Read entire sheet with column names from header row (trimmed)
    Open workbook  ${PROPER EXCEL FILE}  first excel file
    Switch sheet  Sheet 1 (with header)
    @{data sheet}=  Read sheet data  get_column_names_from_header_row=${TRUE}  trim=${TRUE}
    :FOR  ${row}  IN  @{data sheet}
    \  Log dictionary  ${row}
    Close workbook

Read entire sheet with column names from list
    Open workbook  ${PROPER EXCEL FILE}  first excel file
    Switch sheet  Sheet 1 (with header)
    @{column names}=  Create list
    ...  Voornaam
    ...  Achternaam
    ...  Rol
    @{data sheet}=  Read sheet data  column_names=${column names}
    :FOR  ${row}  IN  @{data sheet}
    \  Log dictionary  ${row}
    Close workbook

Read entire sheet without column names
    Open workbook  ${PROPER EXCEL FILE}  first excel file
    Switch sheet  Sheet 1 (with header)
    @{data sheet}=  Read sheet data
    :FOR  ${row}  IN  @{data sheet}
    \  Log list  ${row}
    Close workbook

Read entire sheet with column names from header row
    Open workbook  ${PROPER EXCEL FILE}  first excel file
    Switch sheet  Sheet 1 (with header)
    @{data sheet}=  Read sheet data  cell_range=A1:B3  get_column_names_from_header_row=${TRUE}
    :FOR  ${row}  IN  @{data sheet}
    \  Log dictionary  ${row}
    Close workbook

Read sheet range with column names from list
    Open workbook  ${PROPER EXCEL FILE}  first excel file
    Switch sheet  Sheet 1 (with header)
    @{column names}=  Create list
    ...  Voornaam
    ...  Achternaam
    ...  Rol
    @{data sheet}=  Read sheet data  cell_range=A1:B3  column_names=${column names}
    :FOR  ${row}  IN  @{data sheet}
    \  Log dictionary  ${row}
    Close workbook

Read sheet range without column names (trimmed)
    Open workbook  ${PROPER EXCEL FILE}  first excel file
    Switch sheet  Sheet 1 (with header)
    @{data sheet}=  Read sheet data  cell_range=A1:B3  trim=${TRUE}
    :FOR  ${row}  IN  @{data sheet}
    \  Log list  ${row}
    Close workbook

Writing cell values in temporary sheet with several number formats
    Open workbook  ${PROPER EXCEL FILE}  first excel file
    Run keyword and ignore error
    ...  Remove sheet  TEST
    Create sheet  TEST
    Switch sheet  TEST  # Does NOT implicitly switch when created.
    Write to cell  a1:A1  Hallo
    Write to cell  1, 2  Bart
    Write to cell  coords:1,3  is
    Write to cell  (1,4)  the
    Write to cell  (1, 5)  name
    Write to cell  BBC23  Insurance
    Write to cell  ZZ1  is
    Write to cell  ZZ2  not
    Write to cell  ZZ3  the
    Write to cell  ZZ4  game

    ${A1}=  Read from cell  a1:A1
    ${B1}=  Read from cell  1, 2
    ${C1}=  Read from cell  coords:1,3
    ${D1}=  Read from cell  (1,4)
    ${E1}=  Read from cell  (1, 5)
    ${BBC23}=  Read from cell  BBC23
    ${ZZ1}=  Read from cell  ZZ1
    ${ZZ2}=  Read from cell  ZZ2
    ${ZZ3}=  Read from cell  ZZ3
    ${ZZ4}=  Read from cell  ZZ4

    Should be equal  ${A1}  Hallo
    Should be equal  ${B1}  Bart
    Should be equal  ${C1}  is
    Should be equal  ${D1}  the
    Should be equal  ${E1}  name
    Should be equal  ${BBC23}  Insurance
    Should be equal  ${ZZ1}  is
    Should be equal  ${ZZ2}  not
    Should be equal  ${ZZ3}  the
    Should be equal  ${ZZ4}  game

    Close workbook  alias=first excel file
    # We never saved, so everything should be gone.

Reading and writing with invalid locators
    Open workbook  ${PROPER EXCEL FILE}  first excel file
    Switch sheet  Sheet 1 (with header)
    ${a4}=  Run keyword and expect error  ValueError: * is not a valid column name
    ...  Read from cell  cell=00000000xkaj01xcA4
    Run keyword and expect error  ValueError: * is not a valid column name
    ...  Write to cell  00000000xkaj01xcA4  Hallo




*** Keywords ***
Setup test
    Open workbook   ${PROPER EXCEL FILE}

Teardown test
    Close all workbooks
