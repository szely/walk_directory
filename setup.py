from setuptools import setup

APP = ['main.py']
OPTIONS = {
    'packages': ['openpyxl'],
}

setup(
    app=APP,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)