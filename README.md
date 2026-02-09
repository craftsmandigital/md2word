# md2word

A Python library to inject JSON data into Word templates, automatically converting Markdown strings into formatted Word sub-documents.

## Usage

```python
from md2word import fill_template


data = {
    "title": "# Hello World",
    "content": "This is **bold** text."
}


fill_template("template.docx", data, "output.docx")
