
# Marks Compare — Client-side Dashboard

This project is a client-side Excel comparison tool (Master vs Rovan) built with HTML, CSS, and JavaScript.
It uses SheetJS (XLSX) to parse Excel files in the browser and Chart.js + DataTables for visualization.

## Features
- Upload two Excel files (Master and Rovan)
- Auto-detect orientation (rows or columns)
- Compare marks across Register Nos and Course Codes
- Classify cells: Match / Mismatch / Missing in Master / Missing in Rovan
- Summary cards, per-course chart, detailed table
- Download report as CSV or XLSX
- All processing occurs in the browser (no server)

## How to run locally
1. Open `index.html` in a modern browser (Chrome/Edge/Firefox).
2. Or host on GitHub Pages by pushing repository and enabling Pages (root branch).

## Files
- index.html — main page
- styles.css — styling
- script.js — main logic (SheetJS, DataTables, Chart.js)
