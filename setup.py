"""Contains all the configuration for the package on pip"""
import setuptools

def get_content(*filename:str) -> str:
    """ Gets the content of a file or files and returns
    it/them as a string
    Parameters
    ----------
    filename : (str)
        Name of file or set of files to pull content from 
        (comma delimited)
    
    Returns
    -------
    str:
        Content from the file or files
    """
    content = ""
    for file in filename:
        with open(file, "r") as full_description:
            content += full_description.read()
    return content

setuptools.setup(
    name = "ezspreadsheet",
    version = "0.2.1",
    author = "Kieran Wood",
    author_email = "kieran@canadiancoding.ca",
    description = "A simple API to store/load python objects to/from spreadsheets",
    long_description = get_content("README.md", "CHANGELOG.md"),
    long_description_content_type = "text/markdown",

    project_urls = {
    "API Docs"  :      "https://kieranwood.ca/ezspreadsheet",
    "Source" :         "https://github.com/Descent098/ezspreadsheet",
    "Bug Report":      "https://github.com/Descent098/ezspreadsheet/issues/new?assignees=Descent098&labels=bug&template=bug_report.md&title=%5BBUG%5D",
    "Feature Request": "https://github.com/Descent098/ezspreadsheet/issues/new?assignees=Descent098&labels=enhancement&template=feature_request.md&title=%5BFeature%5D",
    "Roadmap":         "https://github.com/Descent098/ezspreadsheet/projects"
    },
    include_package_data = True,
    py_modules = ["ezspreadsheet"],
    install_requires = [
    "openpyxl", # Used for writing excel files
    "colored",  # Used to colour output for emphasis
        ],
    extras_require = {
        "dev" : ["nox",   # Used to run automated processes
                "pytest", # Used to run the test code in the tests directory
                ],
    },
    classifiers = [
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
        "Development Status :: 4 - Beta"
    ],
)