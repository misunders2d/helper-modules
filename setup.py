from setuptools import setup, find_packages

setup(
    name="helper_modules",           # Name of your package
    version="0.1.0",                    # Version number
    description="A collection of helper modules",
    author="Serge",
    author_email="2djohar@gmail.com",
    packages=find_packages(),           # Automatically find all packages
    install_requires=[                  # Dependencies (if any)
        "customtkinter",
        "tkcalendar"
    ],
    python_requires=">=3.6",            # Specify compatible Python versions
)