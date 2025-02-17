#!/usr/bin/env python3

from setuptools import setup

setup(name='ppt2fig',
      version='1.0.0',
      packages=['ppt2fig'],
      entry_points={
           'console_scripts': [
               'ppt2fig = ppt2fig.main:main'
           ]
      },
      install_requires=[
          'comtypes',
          'pdfCropMargins',
      ]
    )
