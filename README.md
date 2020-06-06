# Certificate Generator - PPTX-PDF Converter
Generate certificate from a powerpoint template. Enter information to the certificate from excel file and generate new certificates in bulk from a single file. And finally turn those generated pptx into pdf files

# How does thos work?

- PPTX_GENERATOR.py script puts the name and id information from name_list.xlsx into certificate_template.pptx powerpoint file row by row.
- In each iteration new pptx file will be generated with the information of iterated row in the excel file.
- PPTX_to_PDF.py will convert all generated pptx files into PDF file.

- To modify the script or template just put a unique text somewhere in the certificate_template.pptx and put those same unique strings into PPTX_GENERATOR.py script.

![certificate Template](README_FILES/template.jpg)
![Name List](README_FILES/names.jpg)
![Generated PPTX](README_FILES/generated_pptx.jpg)
![List of Generated powerpoint files](README_FILES/pptxs.jpg)
![List of Generated PDF files](README_FILES/pdfs.jpg)