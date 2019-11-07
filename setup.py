#!/usr/bin/env python
# -*- coding: utf-8 -*

from __future__ import absolute_import

import os
from codecs import open

from setuptools import find_packages, setup

# allow setup.py to be run from any path
os.chdir(os.path.normpath(os.path.join(os.path.abspath(__file__), os.pardir)))

with open('requirements.txt') as f:
    install_requires = f.read().splitlines()

with open('README.rst', 'r', encoding='utf-8') as rm_file:
    readme = rm_file.read()

with open('HISTORY.rst', 'r', encoding='utf-8') as hist_file:
    history = hist_file.read()

setup(
    name='mega.py',
    version='1.0.2',
    packages=find_packages('src', exclude=('tests', )),
    package_dir={'': 'src'},
    include_package_data=True,
    zip_safe=False,
    description='Python lib for the Mega.co.nz API',
    long_description=readme + '\n\n' + history,
    author='Richard O\'Dwyer',
    author_email='richard@richard.do',
    license='Creative Commons Attribution-Noncommercial-Share Alike license',
    install_requires=install_requires,
    classifiers=[
        'Intended Audience :: Developers',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Topic :: Internet :: WWW/HTTP',
    ]
)
