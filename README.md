import docx
import os
from sys import argv
doc = docx.Document('C:/ Naidu/Downloads/SAHLD - Solution Architecture  High-Level Design (SAHLD)-template-VoLTEspecific-20.8_baselined.docxUsers/Gopi')
all_paras = doc.paragraphs
len(all_paras)

for para in all_paras:
    print(para.text)
    print("-------")








# from docx import Document

import docx
from docx.shared import Inches
mydoc = docx.Document()
mydoc.add_heading("This is level 1 heading", 0)
mydoc.add_paragraph(
    "In untrusted Wi-Fi access, mobility is supported between Wi-Fi access and a co-existing LTE access network. Handovers may occur between the Packet Switched (PS) accesses, based on the common functionality of the evolved packet core. Although the SCC AS of Nokia TAS (MCS profile) and ATCF/ATGW functionality (provided by Nokia SBC) are not required for the PS-PS handover, VoWiFi calls are always anchored in the respective domain transfer anchor in anticipation of a future SRVCC or SRVCC enhanced with ATCF and ATGW event after transfer to LTE. The main characteristics of PS mobility are:")
mydoc.add_paragraph("This is the second paragraph of a MS Word file.")
third_para = mydoc.add_paragraph("This is the third paragraph.")
third_para.add_run(" this is a section at the end of third paragraph")
mydoc.add_heading("adding image into this page", 1)
mydoc.add_picture("F:/unnamed.jpg", width=docx.shared.Inches(5), height=docx.shared.Inches(7))
mydoc.save("E:/needed/dont open/HLD.docx")

mydoc.add_page_break()
# Create an instance of a word document


# Add a Title to the document
mydoc.add_heading('nokia_developers', 0)

# Table data in a form of list
data = (
    (1, 'GOPI 1'),
    (2, 'SUNEEL 2'),
    (3, 'PRAVENN 3')
)

# Creating a table object
table = mydoc.add_table(rows=1, cols=2)

# Adding heading in the 1st row of the table
row = table.rows[0].cells
row[0].text = 'Id'
row[1].text = 'Name'

# Adding data from the list to the table
for id, name in data:
    # Adding a row and then adding data in it.
    row = table.add_row().cells
    # Converting id to string as table can only take string input
    row[0].text = str(id)
    row[1].text = name

# Now save the document to a location
mydoc.save("E:/needed/dont open/HLD.docx")








from docxcompose.composer import Composer
from docx import Document as Document_compose
#filename_master is name of the file you want to merge the docx file into
master = Document_compose('C:/Users/Gopi Naidu/Downloads/SAHLD - Solution Architecture  High-Level Design (SAHLD)-template-VoLTEspecific-20.8_baselined.docx')

composer = Composer(master)
#filename_second_docx is the name of the second docx file
doc2 = Document_compose('E:/needed/dont open/HLD.docx')
#append the doc2 into the master using composer.append function
composer.append(doc2)
#Save the combined docx with a name
composer.save("combineeds.docx")












# importing the modules
import docx
from docx.shared import Inches


class Document_merge():

    def merging_docx(self):
        doc = docx.Document(
            'C:/Users/Gopi Naidu/PycharmProjects/pythonProject5/SAHLD - Solution Architecture  High-Level Design (SAHLD)-template-VoLTEspecific-20.8_baselined (1).docx')
        all_paras = doc.paragraphs
        len(all_paras)

        for para in all_paras:
            print(para.text)
            print("-------")
        #
        # Adding a page break
        doc.add_page_break()
        # add heading into the file
        doc.add_heading("This is level 1 heading", 0)
        doc.add_paragraph(
            "In untrusted Wi-Fi access, mobility is supported between Wi-Fi access and a co-existing LTE access network. Handovers may occur between the Packet Switched (PS) accesses, based on the common functionality of the evolved packet core. Although the SCC AS of Nokia TAS (MCS profile) and ATCF/ATGW functionality (provided by Nokia SBC) are not required for the PS-PS handover, VoWiFi calls are always anchored in the respective domain transfer anchor in anticipation of a future SRVCC or SRVCC enhanced with ATCF and ATGW event after transfer to LTE. The main characteristics of PS mobility are:")

        # adding another heading to the file
        doc.add_paragraph("This is the second paragraph of a MS Word file.")

        third_para = doc.add_paragraph("This is the third paragraph.")
        # aading another
        third_para.add_run(" this is a section at the end of third paragraph")

        doc.add_heading("adding image into this page", 1)
        print('done')
        # adding picture into the file
        doc.add_picture("F:/unnamed.jpg", width=docx.shared.Inches(5), height=docx.shared.Inches(7))


        # adding table into the input file
        doc = docx.Document()
        doc.add_heading('Nokia_product', 0)
        data = (
            (1, 'product 1'),
            (2, 'Cost 2'),
            (3, 'Review 3')
        )
        table = doc.add_table(rows=1, cols=2)
        row = table.rows[0].cells
        row[0].text = 'Id'
        row[1].text = 'Name'
        for id, name in data:
            row = table.add_row().cells
            row[0].text = str(id)
            row[1].text = name
        table.style = 'Colorful List'
        doc.save(
            "gopi_python (2).docx")
        print('done')


merge_obj = Document_merge()

merge_obj.merging_docx()













