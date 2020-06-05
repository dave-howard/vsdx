import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="vsdx",
    version="0.1a1",
    author="Dave Howard",
    author_email="dave@copypy.com",
    description="vsdx - A python library for processing .vsdx files",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/dave-howard/vsdx",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: BSD License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.7',
)