# Task: restructure a policy document

`input/security-policy.docx` is an information security policy with a table of contents, numbered headings ("1 Purpose", "2 Scope", ...), numbered lists and cross-references between sections.

The policy owner wants these changes:

1. **Move** section "6 Data retention", with everything in it (its subsections and table), so that it comes directly after section "3 Access control".
2. **Renumber** the headings so they run 1, 2, 3, ... in the new order again, including the subsection numbers (6.1 and 6.2 become the new section's .1 and .2).
3. The numbered list of steps under "Incident response" currently continues from the list in "Access control" (it starts at 4). Make it **start at 1**.
4. The two **cross-references** in the text ("handled as described in section ...", "are in section ...") must show the referenced heading's new number and title, and their page numbers must be correct.
5. **Update the table of contents** so it matches the new headings and their order.

Change nothing else. Save the result as `work/security-policy.docx`.
