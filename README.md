# Visio2pdf

Convert the first shape of the first page in a Visio file to PDF file.
The converted PDF file will be placed in the same directory of the Visio file.

# Installation

`pip install visio2pdf`

# Usage

## 1. Create Visio2PDFConverter object

```python
from visio2pdf import Visio2PDFConverter

vpc = Visio2PDFConverter(visio_process_name = 'visio.exe',
    current_working_directory = None,
    temp_file_name = 'visio2pdf4latex_temp', visio_ext_names = ['vsdx', 'vsd'])
vpc.convert('one_visio_file.vsd')
```

### `visio_process_name`

The process name of Microsoft visio, default = `visio.exe`. This attribute is used for identifying existing visio process.

### `current_working_directory` {#cwdattribute}

The current working directory, default = `None`. This attribute is used for converting all visio files in the current working directory if no attribute is provided to the `convert` method. Two examples are shown in [Convert all visio files in a folder](#foldercwd) and [Convert all visio files in the current folder](#cwd). 

If `current_working_directory` is `None`, the converter will use the directory determined by the `__file__` attribute provide by the invoking script as current working directory.
If `current_working_directory` is a path, the converter will use the provided path as current working directory. 

### `temp_file_name`

The temporal file name, defalut = `visio2pdf4latex_temp`. During conversion, two temporal files will be created, one `.svg` file and one `.vsdx` file. If `temp_file_name` is set to `WT_`, the temporal files will be `WT_.svg` and `WT_.vsdx`.

### `visio_ext_names` {#extnameattribute}

The extension name of the visio files, default = `['vsdx', 'vsd']`. The converter only converts the files with these extension names.

## 2. Convert visio file to PDF file

### Convert a single `.vsd` file {#file}

This will convert the visio file `one_visio_file.vsd`.

```python
vpc = Visio2PDFConverter()
vpc.convert('one_visio_file.vsd')
```

---

### Convert all visio files with a specified name

This will convert the visio files with a specified name, either `.vsd` or `.vsdx`. 
Both `one_visio_file.vsd` and `one_visio_file.vsdx` will be converted to PDF file.

```python
vpc = Visio2PDFConverter()
vpc.convert('one_visio_file')
```

---

### Convert all visio files in a folder {#folder}

This will convert all visio files in `visio_file_folder_path` and all the sub-directories.
The `visio_file_folder_path` can be provided to the `convert` method.
Or to the [`current_working_directory `](#cwdattribute) when creating the `Visio2PDFConverter`.

```python
vpc = Visio2PDFConverter()
vpc.convert('visio_file_folder_path')
```
or
```python
vpc = Visio2PDFConverter(current_working_directory = 'visio_file_folder_path')
vpc.convert()
```

---

### Convert all visio files in the current working directory {#cwd}

This will convert all visio files in the current folder (determined by `__file__`) and all the sub-directories.

```python
vpc = Visio2PDFConverter()
vpc.convert()
```

---

### Create a converter that only converts `.vsdx` files {#ext}

By default, the created converter will convert all visio files (`.vsd` and `.vsdx`). The extension name can be specified by using [`visio_ext_names`](#extnameattribute) (default value: `['vsdx', 'vsd']`).

```python
vpc = Visio2PDFConverter(visio_ext_names = ['vsdx'])
vpc.convert('one_visio_file.vsdx')
```
