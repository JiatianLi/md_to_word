# markdown to word，to pdf
## build
pyinstaller --noconsole --onefile --icon=icon.ico --name="MarkdownConverter" --add-data "merge_md_to_docx.py;." --add-data "to_pdf.py;." md_converter_gui.py
## install
### xelatex --version
```
https://miktex.org/download
```
### pandoc

