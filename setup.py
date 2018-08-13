import setuptools

setuptools.setup(
    name="robotframework-excellentlibrary",
    version="0.8.4",
    author="Bart Kleijngeld",
    author_email="bartkl@gmail.com",
    description="A really useful Robot Framework library for working with Excel 2010 (and above) files.",
    url="https://github.com/bartkl/robotframework-excellentlibrary",
    packages=setuptools.find_packages(),
    classifiers=(
        "Programming Language :: Python :: 2",
        "Operating System :: OS Independent",
    ),
    python_requires="~=2.7",
    install_requires="openpyxl>=2.5.4",
)
