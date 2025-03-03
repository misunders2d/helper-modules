from setuptools import setup, find_packages

setup(
    name="gui",
    version="0.1",
    packages=["gui"],  # Only includes the gui package
    package_dir={"gui": "gui"},
)