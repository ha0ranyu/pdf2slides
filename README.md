## Project Description

`pdf2slides` is an open-source Python library designed for efficient conversion of PDF files into editable slides. With just three lines of code, it supports both machine-generated PDF files with selectable text and scanned PDF files, although the latter is currently experimental. The library outputs slides in the `.pptx` format, making it a versatile tool for automating presentation creation and handling various types of PDF content. Ideal for integrating into applications or scripts, `pdf2slides` offers a straightforward solution for transforming PDF documents into PowerPoint presentations.

## Installation

To install `pdf2slides`, use pip:

```bash
pip install pdf2slides
```

## Usage

```python
from pdf2slides import Converter

# Create an instance of Converter
converter = Converter()

# Convert a PDF file to slides
converter.convert('input_file.pdf', 'output_file.pptx')
```

## Advanced Usage

You can customize the `Converter` instance with various parameters:

1. **Specifying a Default Font**: Use this parameter to set a font for OCR'ed text. The font must be installed on your system.

```python
from pdf2slides import Converter

converter = Converter(default_font='Arial')
converter.convert('input_file.pdf', 'output_file_with_arial_font.pptx')
```

2. **Enabling OCR Mode**: Enable OCR for converting scanned PDFs. This feature is experimental and is not recommended for use with PDF files known to have selectable text.


```python
from pdf2slides import Converter

converter = Converter(enable_ocr=True)
converter.convert('scanned_file.pdf', 'output_file.pptx')
```

3. **Enforcing Default Font**: Ensure the output slides use the default font, even when the input file is not scanned.

```python
from pdf2slides import Converter

converter = Converter(default_font='Arial', enforce_default_font=True)
converter.convert('not_scanned_file.pdf', 'output_file_with_arial_font.pptx')
```

4. **Manually Setting Image Retention Level**: Adjust the threshold for keeping pure-text images when OCR is enabled.

```python
from pdf2slides import Converter

# Set the image retention level to a lower value when non-editable text remains in the output as images.
converter = Converter(enable_ocr=True, image_retention_level=0.3)
converter.convert('scanned_file.pdf', 'output_file.pptx')
```

5. **Multilingual Support**: Specify the language for OCR processing. The default setting supports English and Chinese. For a list of supported languages, see [PaddleOCR Multi-language Model](https://github.com/PaddlePaddle/PaddleOCR/blob/main/doc/doc_en/multi_languages_en.md).

```python
from pdf2slides import Converter

converter = Converter(enable_ocr=True, lang='fr')
converter.convert('scanned_file_in_french.pdf', 'output_file.pptx')
```

## License

This project is licensed under the GNU General Public License v3.0. See the LICENSE file for details.