# ExcellentLibrary
This library is built on top of OpenPyXL in order to bring its functionality to Robot Framework. The major motivation for this was to add support for Excel 2010 (XSLX) files, which ExcelLibrary does not support. It's important to note that this library does not support the older _xls_ files.

## Installation
ExcellentLibrary can be found on PyPI: https://pypi.org/project/robotframework-excellentlibrary.

To install, simply use pip:

```dos
pip install robotframework-excellentlibrary
```

Dependencies are automatically installed.

## Importing in Robot Framework
As soon as installation has succeeded, you can import the library in Robot Framework:

```robot
*** Settings ***
Library  ExcellentLibrary
```

## Keyword documentation
For the latest keyword documentation [go here](https://bartkl.github.io/robotframework-excellentlibrary/documentation/ExcellentLibrary.html).

If you're using an older version of the library and want to view the appropriate keyword documentation, please open the HTML file locally, or download it from GitHub. It can be found as _ExcellentLibrary.html_ in the _documentation_ directory of the project.
