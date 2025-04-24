import streamlit as st
import zipfile
import tempfile
import os
from PIL import Image
from io import BytesIO
import xlsxwriter

st.set_page_config(page_title="Generador de Excel con Imágenes", layout="centered")

st.title("📸 Generador de Excel con Imágenes")
st.write("Subí un archivo `.zip` que contenga tus imágenes para generar un Excel con vista previa.")

uploaded_file = st.file_uploader("Subí tu archivo ZIP", type=["zip"])

if uploaded_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "images.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_file.read())

        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(tmpdir)

        image_files = [
            os.path.join(tmpdir, f)
            for f in os.listdir(tmpdir)
            if f.lower().endswith((".png", ".jpg", ".jpeg"))
        ]

        if not image_files:
            st.warning("No se encontraron imágenes en el archivo ZIP.")
        else:
            output = BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet()

            # Títulos
            headers = ["IMG", "CODE", "DETAILS", "TOTAL CTNS", "NOTES"]
            for col_num, header in enumerate(headers):
                worksheet.write(0, col_num, header)

            row = 1
            for img_path in image_files:
                img = Image.open(img_path)
                img.thumbnail((150, 150))  # redimensionamos por si son muy grandes

                img_byte_arr = BytesIO()
                img.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)

                worksheet.set_row(row, 220)  # altura de fila

                worksheet.insert_image(row, 0, img_path, {
                    'image_data': img_byte_arr,
                    'x_scale': 1,
                    'y_scale': 1,
                    'x_offset': 5,
                    'y_offset': 5,
                })
                row += 1

            workbook.close()
            output.seek(0)

            st.success("¡Excel generado correctamente!")
            st.download_button(
                label="📥 Descargar Excel",
                data=output,
                file_name="imagenes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
