# Task: turn a Markdown draft into a report on the company template

`input/draft.md` is the draft of the Northwind Water Authority's annual water quality report. `input/northwind-report.dotx` is the company's Word template: its own styles, a header and a footer with page numbers, and a paragraph of guidance text.

Make the report as a Word document based on the template:

- Start from the template, so the document has its styles, header and footer. The template's guidance paragraph must not appear in the report.
- Bring in all of the draft's content: every paragraph, list, the table and the link.
- Styles: the draft's `#` title uses the template's **Report Title** style; `##` headings use **Heading 1**; `###` headings use **Heading 2**; ordinary paragraphs use **Report Body**. Lists stay lists and the table stays a table.
- Put a **table of contents** (Heading 1 and Heading 2) directly after the title and the "Prepared by" line, before the first heading. Its entries and page numbers must be filled in, not empty.
- Set the document's title property to "Annual Water Quality Report 2025" and its author to "Northwind Water Authority".

Save the result as `work/annual-water-quality-report.docx`.
