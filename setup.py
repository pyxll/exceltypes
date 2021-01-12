"""
Excel Types

This package provides type annotations for Excel objects using
the win32com package (or pywin32).

This is also the default type returned by "pyxll.xl_app()" and
so this package will be useful for PyXLL users who want code
completion in editors like PyCharm or Visual Studio Code.

Example usage::

    # _Application is the main "Excel.Application" class
    from exceltypes import _Application

    # If using win32com directly
    import win32com.client
    xl: _Application = win32com.client.Dispatch("Excel.Application")

    # Or if using pyxll
    import pyxll
    xl: _Application = pyxll.xl_app()

"""
from setuptools import setup, find_packages
from os import path


this_directory = path.abspath(path.dirname(__file__))
with open(path.join(this_directory, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()


setup(
    name="exceltypes",
    description="Type annotations for Excel types using win32com / pywin32",
    long_description=long_description,
    long_description_content_type='text/markdown',
    version="0.0.1",
    packages=find_packages(),
    project_urls={
        "Source": "https://github.com/pyxll/exceltypes",
        "Tracker": "https://github.com/pyxll/exceltypes/issues",
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows"
    ],
    python_requires=">=3.8.0",
    install_requires=[
        "pywin32"
    ]
)
