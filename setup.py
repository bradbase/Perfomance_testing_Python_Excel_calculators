"""
FlyingKoala performance tests

"""

import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="xlcalculator_performance_tests_bradbase",
    version="0.0.1b0",
    author="Bradley van Ree",
    author_email="flyingkoala@bradbase.net",
    description="Comparative performance testing of Koala2, PyCel and xlcalculator",
    long_description=long_description,
    long_description_content_type="text/markdown",
    keywords=['xls',
        'excel',
        'spreadsheet',
        'workbook',
        'vba',
        'macro',
        'data analysis',
        'analysis'
        'reading excel',
        'excel formula',
        'excel formulas',
        'excel equations',
        'excel equation',
        'formula',
        'formulas',
        'equation',
        'equations',
        'pandas',
        'harvest',
        'timeseries',
        'time series',
        'energy',
        'accounting',
        'horticulture',
        'research',
        'visualization',
        'scenario analysis',
        'modelling',
        'model',
        'unit testing',
        'testing',
        'audit'],
    url="https://github.com/bradbase/xlcalculator_performance_tests",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
        'Operating System :: Microsoft :: Windows',
        'Operating System :: MacOS :: MacOS X',
    ],
    install_requires=[
            'xlcalculator >= 0.0.1b',
            'pycel >= 1.0b22',
            # 'koala2 >= 0.0.33',
        ]
)
