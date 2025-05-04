---

### 📘 `README.md`

# 🖼️ Image to Spreadsheet

This tool converts an image into a color-coded Excel spreadsheet, where each cell represents a pixel from the image.

## 📂 Project Structure

```
image_to_spreadsheet/
├── imageToSpreadsheet.py      # Main entry point
├── scripts/
│   └── converter.py           # Image conversion logic
├── cat.jpg                    # Example input image
├── README.md                  
```

## ✅ Requirements

* Python 3.7+
* Install dependencies with:

```bash
pip install pillow openpyxl
```

## 🛠️ Usage

Run the main script with:

```bash
python imageToSpreadsheet.py -f cat.jpg -o cat_output.xlsx
```

Or use the optional arguments for width and height with:

```bash
python imageToSpreadsheet.py -f cat.jpg -o cat_output.xlsx --width 80 --height 80
```

### Arguments

| Flag             | Description                                    | Required |
| ---------------- | ---------------------------------------------- | -------- |
| `-f`, `--file`   | Path to the input image file (jpg, png, etc.)  | Yes      |
| `-o`, `--output` | Output Excel file name (must end with `.xlsx`) | Yes      |
| `--width`        | Width to resize the image to (default: 100)    | No       |
| `--height`       | Height to resize the image to (default: 100)   | No       |

## 🔒 Notes

* Supported input formats: `.jpg`, `.jpeg`, `.png`, `.bmp`
* Output format is always `.xlsx` (Excel 2007+)
* Make sure your image isn't too large or the spreadsheet may become slow

## 📷 Example

```bash
python imageToSpreadsheet.py -f cat.jpg -o cat_output.xlsx --width 60 --height 60
```

This will create `cat_output.xlsx` with each cell colored to match the resized image.

---
