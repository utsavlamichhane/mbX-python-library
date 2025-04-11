



from setuptools import setup, find_packages

###Read the README file(my long desc)
with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="pymbX",  #my pypi name 
    version="0.2.0",  #March 7, 2025
    author="Utsav Lamichhane",
    author_email="utsav.lamichhane@gmail.com",
    description="A Python library for downstream pipeline of 16s rRNA sequence data.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    #url="https://github.com/yourusername/mbX",  #will update as soon as paper is published
    packages=find_packages(),  #
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",  
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
    install_requires=[
        "pandas",
        "openpyxl",
        "matplotlib",
        "numpy"
    ],
)
