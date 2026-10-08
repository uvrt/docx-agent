# Task: update a sales memo from new figures

`input/regional-sales-update.docx` is a memo with a sales table, a chart and a few sentences quoting figures. Sales Operations has sent new figures in `input/sales-2026.csv`: they add Q4, and they correct one earlier figure.

Update the memo so that everything agrees with the CSV:

- **The table:** add a Q4 column (between Q3 and Total), correct any changed figure, and recompute the Total column and the Total row.
- **The chart:** add Q4 and correct any changed figure. The chart's own data (what Word shows under Edit Data) must agree with what it draws.
- **The text:** the three figures in the first paragraph (year-to-date total, largest region and its total, strongest quarter and its total) must match the new figures.
- Keep the memo's number format: EUR million with one decimal (`5.9`, `57.1`). Round half up.
- Change nothing else.

Save the result as `work/regional-sales-update.docx`.
