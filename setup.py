#!/usr/bin/env python3

from setuptools import setup

from ppt2fig import __version__

setup(name='ppt2fig',
      version=__version__,
      description='Export PowerPoint pages to PDF with Windows GUIs and a cross-platform CLI / 导出 PowerPoint 页面为 PDF，支持 Windows GUI 和跨平台 CLI。',
      long_description=open('README.md', encoding='utf-8').read(),
      long_description_content_type='text/markdown',
      author='Elliot Zheng',
      author_email='admin@hypercube.top',
      url='https://github.com/elliottzheng/ppt2fig',
      packages=['ppt2fig'],
      entry_points={
           'console_scripts': [
               'ppt2fig = ppt2fig.main:main',
               'ppt2fig-file-gui = ppt2fig.file_gui:main',
           ]
      },
      install_requires=[
          'comtypes; platform_system=="Windows"',
          'pdfCropMargins',
          'pypdf',
      ],
      classifiers=[
          'Programming Language :: Python :: 3',
          'License :: OSI Approved :: MIT License',
          'Operating System :: OS Independent',
      ],
      python_requires='>=3.6'
    )
