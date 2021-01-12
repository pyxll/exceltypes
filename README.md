# Excel Types

This package provides type annotations for Excel objects using
the win32com package (part of pywin32).

This is also the default type returned by "pyxll.xl_app()" and
so this package will be useful for PyXLL users who want code
completion in editors like PyCharm or Visual Studio Code.

Example usage:

    # _Application is the main "Excel.Application" class
    from exceltypes import _Application

    # If using win32com directly
    import win32com.client
    xl: _Application = win32com.client.Dispatch("Excel.Application")

    # Or if using pyxll
    import pyxll
    xl: _Application = pyxll.xl_app()
    
For details of PyXLL please see https://www.pyxll.com.

