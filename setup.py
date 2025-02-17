#!/usr/bin/env python3

from setuptools import setup

setup(name='ppt2fig',
      version='1.0.1',
      description='一键将当前PPT的页面导出为PDF，并自动裁剪白边。',
      long_description=open('README.md', encoding='utf-8').read(),
      long_description_content_type='text/markdown',
      author='Elliot Zheng',
      author_email='admin@hypercube.top',
      url='https://github.com/elliottzheng/ppt2fig',
      packages=['ppt2fig'],
      entry_points={
           'console_scripts': [
               'ppt2fig = ppt2fig.main:main'
           ]
      },
      install_requires=[
          'comtypes',
          'pdfCropMargins',
      ],
      classifiers=[
          'Programming Language :: Python :: 3',
          'License :: OSI Approved :: MIT License',
          'Operating System :: OS Independent',
      ],
      python_requires='>=3.6'
    )
