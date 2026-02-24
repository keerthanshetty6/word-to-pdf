import os
import win32com.client

def convert_folder_to_pdf(folder_path, output_folder=None):
    folder_path = os.path.abspath(folder_path)

    if output_folder:
        output_folder = os.path.abspath(output_folder)
        os.makedirs(output_folder, exist_ok=True)
    else:
        output_folder = folder_path

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False

    try:
        for filename in os.listdir(folder_path):
            if filename.lower().endswith((".docx", ".doc")):
                input_path = os.path.join(folder_path, filename)
                output_filename = os.path.splitext(filename)[0] + ".pdf"
                output_path = os.path.join(output_folder, output_filename)

                print(f"Converting: {filename}")

                doc = word.Documents.Open(input_path)
                doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF
                doc.Close()

        print("✅ Conversion complete.")

    finally:
        word.Quit()


# Example usage
convert_folder_to_pdf(r"C:\Users\Keerthan Shetty\OneDrive - KNIME AG\Desktop\General\vicky\Vicky\word",r"C:\Users\Keerthan Shetty\OneDrive - KNIME AG\Desktop\General\vicky\Vicky\pdf")