import argparse
import os
import sys

sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))
from converter import image_to_excel

def validate_file(path):
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Input file '{path}' does not exist.")
    if not path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
        raise ValueError("Input file must be an image (png, jpg, jpeg, bmp).")

def validate_output(path):
    if not path.strip():
        raise ValueError("Output file name cannot be empty.")
    if not path.lower().endswith('.xlsx'):
        raise ValueError("Output file must have a .xlsx extension (Excel 2007+ format).")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert an image to an Excel file of colored cells.")
    parser.add_argument("-f", "--file", required=True, help="Path to the input image file.")
    parser.add_argument("-o", "--output", required=True, help="Path for the output Excel file.")
    parser.add_argument("--width", type=int, default=100, help="Optional width for the final output inside the spreadsheet.")
    parser.add_argument("--height", type=int, default=100, help="Optional height for the final output inside the spreadsheet.")

    args = parser.parse_args()

    try:
        validate_file(args.file)
        validate_output(args.output)
        resize_to = (args.width, args.height)
        output_filename = image_to_excel(args.file, args.output, resize_to)
        print(f"Saved resized image to spreadsheet: {output_filename}")
    except Exception as e:
        print(f"Error: {e}")
