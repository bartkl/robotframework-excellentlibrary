import setuptools

setuptools.setup(
    name="robotframework-excellentlibrary",
    version="1.0.0",
    author="Bart Kleijngeld",
    author_email="bartkl@gmail.com",
    description="A really useful Robot Framework library for working with Excel 2010 (and above) files.",
    url="https://github.com/bartkl/robotframework-excellentlibrary",
    packages=setuptools.find_packages(),
    classifiers=(
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ),
    python_requires="~=3.6",
    install_requires="openpyxl>=3.0.3",
    keywords="excel testing robotframework robotframework-excellibrary robotframework-excellib"
)
