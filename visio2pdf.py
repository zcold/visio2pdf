# coding = utf-8
# The MIT License (MIT)
  # Copyright (c) 2015 by Shuo Li (contact@shuo.li)
  #
  # Permission is hereby granted, free of charge, to any person obtaining a
  # copy of this software and associated documentation files (the "Software"),
  # to deal in the Software without restriction, including without limitation
  # the rights to use, copy, modify, merge, publish, distribute, sublicense,
  # and/or sell copies of the Software, and to permit persons to whom the
  # Software is furnished to do so, subject to the following conditions:
  #
  # The above copyright notice and this permission notice shall be included in
  # all copies or substantial portions of the Software.
  #
  # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  # FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  # AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  # LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
  # FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
  # DEALINGS IN THE SOFTWARE.

'''Convert the first shape of the first page in a Visio file to PDF file.'''

__author__ = 'Shuo Li <contact@shuol.li>'
__version__= '2015-08-04-09:22'

import os, os.path, glob, shutil, argparse, logging, warnings, time
import psutil, win32com.client

def process_exists(process_name) :

  '''Exam if a process with the given name is exists.'''

  for p in process_iter():
    try :
      if p.name().lower() == process_name :
        return True
    except :
      pass
  return False

def parse_command_line_args(current_file = __file__,
  visio_ext_names = ['vsdx', 'vsd']) :

  cwd = os.path.dirname(os.path.abspath(current_file))
  parser = argparse.ArgumentParser(description = 'visio to pdf converter')
  parser.add_argument('visio_files',
    nargs = '*', help = 'visio files',
    default = [] )
  parser.add_argument('--more_visio_files', '-f',
    nargs = '*', help = 'visio files',
    default = [] )
  parser.add_argument('--temp_file_name', '-t',
    help = 'temporal file name',
    default = 'visio2pdf4latex_temp' )
  parser.add_argument('--ext_names', '-e',
    nargs = '*', help = 'visio file extension names',
    default = ['vsdx', 'vsd'] )

  return parser.parse_args()

class Visio2PDFConverter(object) :

  def __init__(self, visio_process_name = 'visio.exe',
    current_working_directory = None,
    temp_file_name = 'visio2pdf4latex_temp', visio_ext_names = ['vsdx', 'vsd']) :

    self.visio_ext_names = visio_ext_names
    self.cwd = self.__get_cwd(current_working_directory)
    self.visio = self.__new_visio_process(visio_process_name)
    self.svg_temp_file_path = os.path.join(self.cwd, temp_file_name + '.svg')
    self.vsdx_temp_file_path = os.path.join(self.cwd, temp_file_name + '.vsdx')
    self.__remove_temporal_files()

  def __del__(self) :
    self.visio.Quit()

  def __get_cwd(self, provided_working_directory) :
    dir_or_file = provided_working_directory or __file__
    return os.path.dirname(os.path.abspath(dir_or_file))

  def __new_visio_process(self, visio_process_name) :
    self.visio_process_name = visio_process_name
    visio = win32com.client.Dispatch('Visio.Application')
    visio.Visible = 0
    return visio

  def __remove_temporal_files(self) :
    for f in [self.svg_temp_file_path, self.vsdx_temp_file_path] :
      if os.path.isfile(f) :
        os.remove(f)

  def convert_one_file(self, visio_file_path) :

    print('converting ' + visio_file_path)

    # make a copy of the visio file
    shutil.copyfile(visio_file_path, self.vsdx_temp_file_path)

    # save the 1st shape in the 1st page to temp svg
    # otherwise the pdf will be A4 or other paper size
    doc = self.visio.Documents.Open(FileName = self.vsdx_temp_file_path)
    #doc.Pages[0].Shapes[0].Export(self.svg_temp_file_path)
    doc.Pages[0].Export(self.svg_temp_file_path) # considering export all element in page. 
    doc.Close()

    # convert temp.svg into pdf file
    doc = self.visio.Documents.Open(FileName = self.svg_temp_file_path)
    pdf_file_name = os.path.splitext(visio_file_path)[0] + '.pdf'
    pdf_file_path = os.path.join(self.cwd, pdf_file_name)
    doc.ExportAsFixedFormat(1, pdf_file_path,
      0, 0,
      IncludeBackground = True,
      IncludeDocumentProperties = False,
      UseISO19005_1 = True)
    doc.Close()

    print('done ')

  def visio_files_in_path(self, file_path):

    visio_files = []
    if os.path.isdir(file_path) :
      for ext_name in self.visio_ext_names :
        visio_files += glob.glob(os.path.join(file_path, '*.' + ext_name))
    elif os.path.isfile(file_path) :
      if file_path.lower().endswith('vsdx') or file_path.lower().endswith('vsd') :
        visio_files.append(file_path)
      else :
        warn_msg = '{0} is not a Visio file name.'.format(f)
        warnings.warn(warn_msg)
    else :
        warn_msg = '{} does not exist.'.format(f)
        warnings.warn(warn_msg)
    return visio_files

  def convert(self, files = None) :

    visio_files = []
    files = files or []
    if isinstance(files, str) :
      files = [files]

    if len(files) == 0 :
      for ext_name in self.visio_ext_names :
        visio_files += glob.glob(os.path.join(self.cwd, '*.' + ext_name))
    else :
      for f in files :
        file_path = os.path.abspath(f)
        visio_files += self.visio_files_in_path(file_path)

    for visio_file in visio_files:
      self.convert_one_file(visio_file)
      self.__remove_temporal_files()

if __name__ == '__main__' :

  args = parse_command_line_args()
  vpc = Visio2PDFConverter(temp_file_name = args.temp_file_name,
    visio_ext_names = args.ext_names)
  vpc.convert(args.visio_files + args.more_visio_files)

