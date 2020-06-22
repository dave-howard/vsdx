import setuptools

# usage notes - to prepare and upload to PyPI:
# if missing, install setuptools, wheel and twine using pip
# delete /build, /dist and /vdx.egg-info directories
# >>> python setup.py sdist bdist_wheel
# >>> python -m twine upload dist/*

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="vsdx",
    version="0.2",
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