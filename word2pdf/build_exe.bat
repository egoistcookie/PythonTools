@echo off

pip install docx2pdf
pip install pyinstaller

pyinstaller --onefile --windowed --noconsole --name "WordToPdfConverter_new" word_to_pdf_converter.py

echo 打包完成！