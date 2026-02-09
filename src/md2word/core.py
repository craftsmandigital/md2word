import os
import tempfile
import uuid
import re
from docxtpl import DocxTemplate
import pypandoc

# --- MONKEY PATCH START ---
# Fix for docxcompose on newer Python versions
try:
    import pkg_resources
except ImportError:
    pass
# --- MONKEY PATCH END ---

def _is_markdown(text):
    """Check if text contains Markdown syntax."""
    if not isinstance(text, str):
        return False
    # Regex for Headers, Bold, Italic, Lists, Links
    pattern = re.compile(
        r"(\*{1,2}[^*\n]+\*{1,2}|_{1,2}[^_\n]+_{1,2}|#{1,6}\s|\[.*?\]\(.*?\)|^\s*[-*+]\s|^\s*\d+\.\s)",
        re.MULTILINE
    )
    return '\n' in text or bool(pattern.search(text))

def _process_value(value, doc, temp_dir):
    """
    Recursive helper to process dictionary/list/strings.
    """
    if isinstance(value, dict):
        return {k: _process_value(v, doc, temp_dir) for k, v in value.items()}
    
    elif isinstance(value, list):
        return [_process_value(item, doc, temp_dir) for item in value]
    
    elif isinstance(value, str) and _is_markdown(value):
        # Convert Markdown -> Docx Subdocument
        unique_name = f"{uuid.uuid4()}.docx"
        temp_path = os.path.join(temp_dir, unique_name)
        
        try:
            pypandoc.convert_text(
                value, 
                'docx', 
                format='markdown', 
                outputfile=temp_path,
                extra_args=['--from=markdown+hard_line_breaks']
            )
            if os.path.exists(temp_path) and os.path.getsize(temp_path) > 0:
                return doc.new_subdoc(temp_path)
        except Exception as e:
            print(f"Warning: Failed to convert markdown snippet. Error: {e}")
        
        return value # Fallback to raw text
    
    else:
        return value

def fill_template(template_path, data, output_path):
    """
    Load a template, inject data (converting Markdown), and save.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    doc = DocxTemplate(template_path)


    # Use a temporary directory for the intermediate sub-documents
    with tempfile.TemporaryDirectory() as temp_dir:
        # Recursively process the data dictionary
        context = _process_value(data, doc, temp_dir)
        
        # Render the template
        doc.render(context)
        
        # Save the result
        doc.save(output_path)
        print(f"Document saved to: {output_path}")
