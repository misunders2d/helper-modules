from setuptools import setup, find_packages

setup(
    name="helper_modules",  # Name of your package
    version="0.1.43",  # Version number
    description="A collection of helper modules",
    author="Sergey",
    author_email="2djohar@gmail.com",
    packages=find_packages(),  # Automatically find all packages
    install_requires=[  # Dependencies (if any)
        "customtkinter",
        "gdown",
        "google-api-python-client",
        "google-cloud-bigquery",
        "gspread",
        "numpy",
        "openpyxl",
        "python-dotenv",
        "pandas",
        "pandas-gbq",
        "tkcalendar",
        "xlsxwriter",
    ],
    python_requires=">=3.6",  # Specify compatible Python versions
)
