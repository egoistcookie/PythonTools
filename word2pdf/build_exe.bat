@echo off

pip install docx2pdf==0.1.8
pip install pyinstaller==5.13.0

pyinstaller --onefile --windowed --noconsole --name "WordToPdfConverter" word_to_pdf_converter.py

echo 打包完成！