# Sticker Printer Tool

A fully customizable web-based application designed to precisely generate and print price tags and barcode labels on standard A4 sticker sheets. 

---

**page link:** ([apple id generator](https://lian-kanani.github.io/Sticker-Printer-Tool/))

## Features
* **Precise Measurements:** Custom inputs for top/left margins, label height/width, and vertical/horizontal steps (in centimeters) to fit perfectly with any pre-cut sticker sheet.
* **Skip Used Labels:** Input the number of already peeled stickers on your current sheet, and the program will automatically skip those spaces to save paper.
* **Excel Integration:** Upload `.xlsx` files to bulk-import items (Name, Fonts, Price, Currency, Quantity).
* **Manual Entry Table:** A dynamic table to add, edit, or delete items on the fly with custom font sizes for names and prices.
* **Print-Ready A4 Generation:** Automatically distributes labels across multiple A4 pages with exact CSS grid calculations.

## How to Use

1. Clone the repository: 
```Bash
git clone https://github.com/lian-kanani/Sticker-Printer-Tool.git
cd Sticker-Printer-Tool
```
2. Open index.html in your web browser (Google Chrome recommended).

3. Adjust the Layout Settings (إعدادات التخطيط) to match your physical sticker sheet.

4. If using a partially used sheet, enter the number of missing labels in the Skip Labels section.

5. Import your data via Excel OR add them manually in the table.

6. Click Preview (معاينة) to generate the A4 sheets.

7. Click Print Now (طباعة الآن) and ensure your printer is set to A4 size, Default Margins, and Actual Size (100% scale).

## Tech Stack

- HTML5 / CSS3 (Grid & Print Media Queries)
- Vanilla JavaScript
- [SheetJS (xlsx)](https://sheetjs.com/) for Excel parsing.
