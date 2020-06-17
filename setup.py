"""
Description:
    Contains all the configuration for the package on pip
"""
import setuptools

def get_content(*filename):
    """ Gets the content of a file and returns it as a string
    Args:
        filename(str): Name of file to pull content from
    Returns:
        str: Content from file
    """
    content = ""
    for file in filename:
        with open(file, "r") as full_description:
            content += full_description.read()
    return content

setuptools.setup(
    name = "ezexcel",
    version = "0.0.1", # I recommend every 2nd decimal release for big releases and 3rd for bug fixes.
    author = "Kieran Wood",
    author_email = "kieran@canadiancoding.ca",
    description = "A simple class based xlsx serialization system",
    long_description = get_content("README.md", "CHANGELOG.md"),
    long_description_content_type = "text/markdown",
    url = "https://github.com/Descent098/ezexcel",
    include_package_data = True,
    py_modules = ["ezexcel"],
    

    install_requires = [
    "XlsxWriter", # Used for writing excel files
      ],
    extras_require = {
        "dev" : ["nox",    # Used to run automated processes
                "pytest", # Used to run the test code in the tests directory
                ],

    },
    classifiers = [
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
)