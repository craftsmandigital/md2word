# md2word

**Inject JSON data into Word templates with automatic Markdown formatting.**

`md2word` is a Python library that takes a Python dictionary (or JSON), detects Markdown syntax (like `**bold**`, lists, or headers), and converts those specific fields into real, formatted Word sub-documents inside your master template.

It bridges the gap between raw data and beautifully formatted Word reports.

---

## ⚠️ Compatibility Warning

**Python Version:** Supports Python **3.9** through **3.12**.
**Python 3.13 is currently NOT supported.**

This library relies on `docxcompose` and older `setuptools` features (`pkg_resources`) that were removed in Python 3.13. If you are using `uv`, ensure you pin your python version:

```bash
uv python pin 3.12
```

---

## 🚀 Installation

This package is hosted on GitHub. You can install it directly using `uv` (recommended) or pip.

### Using uv

```bash
uv add "md2word @ git+https://github.com/craftsmandigital/md2word.git"
```

### Using pip

```bash
pip install "git+https://github.com/craftsmandigital/md2word.git"
```

*Note: This package automatically installs `pypandoc-binary`, so you do not need to install Pandoc manually on your system.*

---

## 📝 Usage

### 1. Python Code

```python
from md2word import fill_template

# 1. Define your data
# Strings with Markdown syntax are automatically detected and converted.
data = {
    "project_name": "Project Phoenix",
    "date": "2026-02-10",
    "summary": "# Executive Summary\n\nThe project is **on track**. We have completed:\n- Phase 1\n- Phase 2",
    "risks": [
        {"title": "Budget", "details": "Risk is **High**. Mitigation: *Cut costs*."},
        {"title": "Time", "details": "Risk is Low."}
    ]
}

# 2. Generate the document
# fill_template(template_path, data_dict, output_path)
fill_template("template.docx", data, "report.docx")
```

### 2. Word Template Syntax

This library uses **[docxtpl](https://docxtpl.readthedocs.io/)**, which relies on the **[Jinja2](https://jinja.palletsprojects.com/)** templating language.

Open your `.docx` file and use these placeholders:

#### Simple Variables

Use double curly braces:

```text
Project: {{ project_name }}
Date: {{ date }}
```

#### Markdown Content

Just use the variable name. The library will automatically swap the text for a formatted Word block if it detects Markdown.

```text
{{ summary }}
```

#### Loops (Tables and Lists)

You can loop through arrays using `{% for %}` tags. This is perfect for Word tables.

| Risk Title | Details |
| :--- | :--- |
| `{% for r in risks %}` {{ r.title }} | {{ r.details }} `{% endfor %}` |

> **💡 Important Tip for Loops:**
> When rendering Markdown inside a loop (like a table cell), ensure there is a **space or new line** between the variable and the end tag.
>
> * ❌ **Bad:** `{{ r.details }}{% endfor %}` (Might crash Word)
> * ✅ **Good:** `{{ r.details }} {% endfor %}`

---

## 📂 Examples

A complete working example is included in this repository.

Check the **[examples/](examples/)** folder to see:

1. **`demo.py`**: A script showing how to structure data.
2. **`template_m2word.docx`**: A reference Word template showing how to set up headers, tables, and loops.

To run the example locally:

```bash
# 1. Clone the repo
git clone https://github.com/craftsmandigital/md2word.git
cd md2word

# 2. Run the demo script
uv run examples/demo.py
```

---

## ✅ Supported Markdown

The following Markdown syntax is converted into native Word formatting:

* **Bold** (`**text**`)
* *Italic* (`*text*` or `_text_`)

* # Headers (`# H1`, `## H2`, etc.)

* [Links](https://example.com) (`[Text](url)`)
* Lists (Bulleted `- item` and Numbered `1. item`)

---

## 🛠 Dependencies

This package stands on the shoulders of giants:

* [docxtpl](https://github.com/elapouya/python-docx-template) - For Jinja2 templating in Word.
* [pypandoc](https://github.com/JessicaTegner/pypandoc) - For converting Markdown to Docx.
* [docxcompose](https://github.com/4teamwork/docxcompose) - For merging the converted Markdown sub-documents.
