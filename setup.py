import setuptools

# usage notes - to prepare and upload to PyPI:
# if missing, install setuptools, wheel and twine using pip
# delete /build, /dist and /vdx.egg-info directories
# >>> python setup.py sdist bdist_wheel
# >>> python -m twine upload dist/*
# note: use __token__ as user and actual token as password

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="vsdx",
    version="0.5.4",
    author="Dave Howard",
    author_email="dave@codypy.com",
    description="vsdx - A python library for processing .vsdx files",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/dave-howard/vsdx",
    packages=setuptools.find_packages(),
    install_requires=[
        'Jinja2',
        'deprecation',
    ],
    classifiers=[
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "License :: OSI Approved :: BSD License",
        "Operating System :: OS Independent",
        "Development Status :: 4 - Beta",
    ],
    project_urls={
        'Documentation': 'https://vsdx.readthedocs.io/en/latest/'
    },
    python_requires='>=3.7',
    package_data={
        "": ["media/*.vsdx"],
    },
)
