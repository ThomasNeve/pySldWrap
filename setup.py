from setuptools import setup
import pathlib

here = pathlib.Path(__file__).parent.resolve()
long_description = (here / 'README.md').read_text(encoding='utf-8')

setup(
  name = 'pySldWrap',
  packages = ['pySldWrap'],
  version = '0.1',
  license='MIT',        # Chose a license from here: https://help.github.com/articles/licensing-a-repository
  description = 'Python Solidworks interface',
  long_description = long_description,
  author = 'Thomas Neve',
  author_email = 'thomas.neve@ugent.be',
  url = 'https://github.com/ThomasNeve/pySldWrap',
  download_url = 'https://github.com/ThomasNeve/pySldWrap/archive/v_01.tar.gz',
  keywords = ['solidworks', 'wrapper'],
  install_requires=[            # I get to this in a second
          'numpy',
          'pywin32',
          'pathlib'
      ],
  classifiers=[
    'License :: OSI Approved :: MIT License',
    'Programming Language :: Python :: 3.5',
    'Programming Language :: Python :: 3.6',
    'Programming Language :: Python :: 3.7'
  ],
  python_requires='>=3.6',
)