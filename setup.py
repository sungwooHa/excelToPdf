from setuptools import setup, find_packages

setup(
    name="excel-to-pdf",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "pywin32>=228",
        "packaging",
    ],
    entry_points={
        "console_scripts": [
            "excel-to-pdf=excel_to_pdf:main",
            "excel-to-pdf-gui=excel_to_pdf_gui:main",
        ],
    },
    python_requires=">=3.6",
    description="Convert Excel files to PDF format",
    author="Your Name",
    author_email="your.email@example.com",
    url="https://github.com/yourusername/excelToPdf",
)