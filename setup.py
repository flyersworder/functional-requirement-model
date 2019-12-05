#!/usr/bin/env python

from setuptools import setup, find_packages

setup(
    name='frm_model',
    version='0.0.1',
    url='https://bitbucket.org/zhaidewei/frm/src/master/',
    packages=find_packages('src'),
    package_dir={'': 'src'},
    include_package_data=True,
    python_requires='>=3.7.4',
    install_requires=[
        'pyyaml',
        'pandas',
        'numpy',
        'fuzzywuzzy',
        'pulp',
        'plotly',
        'python-Levenshtein',
        'openpyxl',
        'streamlit'
    ],
    extras_require={
        'test': [
            'pylint',
            'pytest'
        ]
    },
    zip_safe=True,
)
