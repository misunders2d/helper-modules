from setuptools import setup, find_packages

setup(
    name="ctk_gui",
    version="0.1",
    packages=find_packages(),
    package_dir={'': '.'}  # Important if not at root
)