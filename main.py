from PIL import Image
import os

# Excel conversion (Windows only)
try:
    import win32com.client
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


def convert_tiff_to_pdf(input_path, output_path):
    with Image.open(input_path) as img:
        if getattr(img, "n_frames", 1) > 1:
            images = []
            for i in range(img.n_frames):
                img.seek(i)
                images.append(img.convert("RGB"))
            images[0].save(output_path, save_all=True, append_images=images[1:])
        else:
            img.convert("RGB").save(output_path, "PDF")


def convert_excel_to_pdf(input_path, output_path):
    if not EXCEL_AVAILABLE:
        print(f"Skipping Excel file (missing pywin32): {input_path}")
        return

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        wb = excel.Workbooks.Open(os.path.abspath(input_path))
        wb.ExportAsFixedFormat(0, os.path.abspath(output_path))
        wb.Close(False)
        excel.Quit()

    except Exception as e:
        print(f"Error converting Excel {input_path}: {e}")


def convert_all_to_pdf(input_folder="input", output_folder="pdf"):
    """
    Convert TIFF and Excel files in input_folder to PDF in output_folder.
    Supports:
    - TIFF/TIF
    - XLS, XLSX, XLSM
    """

    os.makedirs(output_folder, exist_ok=True)

    files = os.listdir(input_folder)

    if not files:
        print(f"No files found in '{input_folder}' folder.")
        return

    for file in files:
        input_path = os.path.join(input_folder, file)
        filename, ext = os.path.splitext(file)
        ext = ext.lower()

        output_path = os.path.join(output_folder, filename + ".pdf")

        try:
            if ext in [".tif", ".tiff"]:
                convert_tiff_to_pdf(input_path, output_path)
                print(f"Converted TIFF: {file} → {filename}.pdf")

            elif ext in [".xls", ".xlsx", ".xlsm"]:
                convert_excel_to_pdf(input_path, output_path)
                print(f"Converted Excel: {file} → {filename}.pdf")

            else:
                print(f"Skipped (unsupported): {file}")

        except Exception as e:
            print(f"Error converting {file}: {e}")

    print(f"\nAll conversions completed! PDFs saved in '{output_folder}/'.")


# Run
if __name__ == "__main__":
    convert_all_to_pdf("input", "pdf")