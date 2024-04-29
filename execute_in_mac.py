from setuptools import setup
import py2app

APP = ["from_excel_to_pdf_rafa.py"]
OPTIONS = {
    "argv_emulation": True,
}

setup(
    app=APP,
    options={
        'py2app': OPTIONS
    },
    setup_requires=["py2app"]
)
