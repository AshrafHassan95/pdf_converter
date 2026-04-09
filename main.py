from PIL import Image
import os
import time

# Excel conversion (Windows only)
try:
    import win32com.client
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


# =========================
# TIFF → PDF
# =========================
def convert_tiff_to_pdf(input_path, output_path):
    try:
        with Image.open(input_path) as img:
            if getattr(img, "n_frames", 1) > 1:
                images = []
                for i in range(img.n_frames):
                    img.seek(i)
                    images.append(img.convert("RGB"))
                images[0].save(output_path, save_all=True, append_images=images[1:])
            else:
                img.convert("RGB").save(output_path, "PDF")

        return True

    except Exception as e:
        print(f"❌ TIFF Error: {input_path} → {e}")
        return False


# =========================
# Excel → PDF (Stable Version)
# =========================
def convert_excel_to_pdf(input_path, output_path, retries=3):
    if not EXCEL_AVAILABLE:
        print(f"⚠️ Skipped Excel (pywin32 not installed): {input_path}")
        return False

    for attempt in range(retries):
        excel = None

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False  # 🔥 disable popups

            wb = excel.Workbooks.Open(os.path.abspath(input_path))

            # Optional: ensure printable
            for sheet in wb.Worksheets:
                sheet.PageSetup.Zoom = False

            wb.ExportAsFixedFormat(0, os.path.abspath(output_path))
            wb.Close(False)
            excel.Quit()

            return True  # ✅ success

        except Exception as e:
            print(f"⚠️ Retry {attempt + 1}/{retries} for Excel: {os.path.basename(input_path)}")

            if excel:
                try:
                    excel.Quit()
                except:
                    pass

            time.sleep(2)

            if attempt == retries - 1:
                print(f"❌ Excel Failed: {input_path} → {e}")
                return False


# =========================
# MAIN BATCH PROCESS
# =========================
def convert_all_to_pdf(input_folder="input", output_folder="pdf"):
    os.makedirs(output_folder, exist_ok=True)

    files = os.listdir(input_folder)

    if not files:
        print(f"⚠️ No files found in '{input_folder}'")
        return

    success_count = 0
    fail_count = 0
    skip_count = 0

    for file in files:
        input_path = os.path.join(input_folder, file)
        filename, ext = os.path.splitext(file)
        ext = ext.lower()

        output_path = os.path.join(output_folder, filename + ".pdf")

        # Skip system files
        if file.lower() == "desktop.ini":
            continue

        print(f"\n🔄 Processing: {file}")

        success = False

        try:
            if ext in [".tif", ".tiff"]:
                success = convert_tiff_to_pdf(input_path, output_path)

            elif ext in [".xls", ".xlsx", ".xlsm"]:
                success = convert_excel_to_pdf(input_path, output_path)

            else:
                print(f"⏭️ Skipped (unsupported): {file}")
                skip_count += 1
                continue

            if success:
                print(f"✅ Converted: {file} → {filename}.pdf")
                success_count += 1
            else:
                fail_count += 1

        except Exception as e:
            print(f"❌ Unexpected error: {file} → {e}")
            fail_count += 1

        time.sleep(1)  # 🔥 prevent Excel overload

    # =========================
    # SUMMARY
    # =========================
    print("\n=========================")
    print("📊 Conversion Summary")
    print("=========================")
    print(f"✅ Success : {success_count}")
    print(f"❌ Failed  : {fail_count}")
    print(f"⏭️ Skipped : {skip_count}")
    print(f"\n📁 Output folder: '{output_folder}/'")


# =========================
# RUN
# =========================
if __name__ == "__main__":
    convert_all_to_pdf("input", "pdf")