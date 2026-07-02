"""Deterministic DOCX / XLSX engines.

The PPTX pipeline lives in ``core/`` (inherited from edit2ppt); this
package adds the same deterministic building blocks for the other two
OOXML families:

* ``docx_engine`` — markdown -> Word, paragraph outline/addressing,
  targeted text edits, Word -> markdown.
* ``xlsx_engine`` — sheet spec -> Excel, workbook outline, targeted cell
  edits, Excel -> markdown tables.

LLM orchestration on top of these lives in ``tools/generate_doc.py`` and
``tools/edit_doc.py``.
"""
