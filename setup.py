# coding = utf-8
# The MIT License (MIT)
  # Copyright (c) 2015 by Shuo Li (contact@shuo.li)
  #
  # Permission is hereby granted, free of charge, to any person obtaining a
  # copy of this software and associated documentation files (the 'Software'),
  # to deal in the Software without restriction, including without limitation
  # the rights to use, copy, modify, merge, publish, distribute, sublicense,
  # and/or sell copies of the Software, and to permit persons to whom the
  # Software is furnished to do so, subject to the following conditions:
  #
  # The above copyright notice and this permission notice shall be included in
  # all copies or substantial portions of the Software.
  #
  # THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
  # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
  # DEALINGS IN THE SOFTWARE.

__author__ = 'Shuo Li <contact@shuol.li>'
__version__= '2015-08-12-08:12'

from setuptools import setup

setup(
  name = 'visio2pdf',
  version = '1.0.0',
  description = 'Convert the first shape of the first page of a Visio file to cropped PDF file.',
  long_description = 'Convert the first shape of the first page of a Visio file to cropped PDF file. Microsoft Visio SHOULD be installed.',
  install_requires = ['psutil>=3.1.1', 'pywin32>=219'],
  author = 'Shuo Li',
  author_email = 'contact@shuo.li',
  maintainer = 'Shuo Li',
  maintainer_email = 'contact@shuo.li',
  url = 'https://github.com/zcold/visio2pdf',
  download_url = 'https://github.com/zcold/visio2pdf/tarball/1.0.0',
  py_modules = ['visio2pdf'],
  license = 'MIT',
  keywords = ['Visio', 'PDF'],
  platforms = ['win32'],
  classifiers = [
    'Development Status :: 4 - Beta',
    'Topic :: Multimedia :: Graphics',
    'License :: OSI Approved :: MIT License']
)
