# PDF Compression and Merging Tool

This project provides a tool for compressing PDF files by reducing image quality and resizing images inside the PDF. It also includes functionality for merging multiple PDF files into a single PDF. This tool is designed to run on Windows.

## Features
- **Compress PDF files** by reducing the resolution of images and adjusting image quality.
- **Merge multiple PDFs** into a single PDF file.
- Supports both Ghostscript and pikepdf for PDF processing.

## Requirements

### Python Dependencies
Before running the project, make sure to install the following Python dependencies:

1. **PyPDF2**: A library used for merging PDFs.
   ```sh
   pip install PyPDF2
   ```
2. **pikepdf**: A library for manipulating PDF files (including compressing and resizing images).
   ```sh
   pip install pikepdf
   ```
3. **Pillow**: A Python Imaging Library (PIL) fork to handle images within PDFs.
   ```sh
   pip install Pillow
   ```

### Ghostscript
This project uses **Ghostscript** to compress images within PDF files. You need to download and install it on your system.

1. Download Ghostscript from [https://ghostscript.com/releases/gsdnld.html](https://ghostscript.com/releases/gsdnld.html).
2. Follow the installation instructions provided on the Ghostscript website.

### Verify Ghostscript Installation
After installing Ghostscript, verify that it is correctly installed by running the following command in the Command Prompt:

```sh
gswin64c --version
```

If Ghostscript is installed correctly, you should see the version information.

## Usage

1. Place the PDFs you want to compress and/or merge into a folder.
2. Run the Python script by specifying the input and output directories as follows:
   ```sh
   python PDF-Optimizer.py
   ```

## License
This project is licensed under the MIT License.

