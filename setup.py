import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="excel",
    version="0.0.1",
    author="neo zhang",
    author_email="suganzhang123@gmail.com",
    description="A simple package for excel options",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Wowgreat/excel_tools",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)