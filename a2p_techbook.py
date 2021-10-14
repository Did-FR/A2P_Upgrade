#***************************************************************************
#*                                                                         *
#*   Copyright (c) 2021 didier                                             *
#*                                                                         *
#*   This program is free software; you can redistribute it and/or modify  *
#*   it under the terms of the GNU Lesser General Public License (LGPL)    *
#*   as published by the Free Software Foundation; either version 2 of     *
#*   the License, or (at your option) any later version.                   *
#*   for detail see the LICENCE text file.                                 *
#*                                                                         *
#*   This program is distributed in the hope that it will be useful,       *
#*   but WITHOUT ANY WARRANTY; without even the implied warranty of        *
#*   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the         *
#*   GNU Library General Public License for more details.                  *
#*                                                                         *
#*   You should have received a copy of the GNU Library General Public     *
#*   License along with this program; if not, write to the Free Software   *
#*   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  *
#*   USA                                                                   *
#*                                                                         *
#***************************************************************************

import os
import datetime
from pathlib import Path
from PyPDF2 import PdfFileMerger
from PySide import QtGui, QtCore
import FreeCADGui
import FreeCAD
import Spreadsheet
import TechDraw
import TechDrawGui

from a2p_translateUtils import *
import a2plib
from a2p_simpleXMLreader import FCdocumentReader

toolTip = translate("A2plus",
"""
Create a pdf file containing
tech draw export included in
the assembly.

Editable fields of the templates
can be set thanks ti the 
#TECHINFO# spreadsheet.
The fields date, scale and sheet
will be calculated.

This button will open a dialog
with the Question:
- Iterate recursively over all subassemblies?

Answer Yes:
All parts of all subassemblies are
collected to the tech book.

Answer No:
Only the parts within the
recent assembly are collected.
"""
)


class A2PCreateTechBook:
    STANDARDS_FIELDS_NAME = ("Stating_Page", "Nb_Page_After", "Date_Field", "Scale_Field", "Sheet_Field")
    INFO_SPREADSHEET_LABEL = "#TECHINFO#"

    def _getUserParameters(self, working_dir):
        """ Asks questions to user and return information given.
        :return: (recursive_call, export_filename) or None
        """
        flags = QtGui.QMessageBox.StandardButton.Yes | QtGui.QMessageBox.StandardButton.No
        msg = u"Do you want to iterate recursively over all included subassemblies?"
        response = QtGui.QMessageBox.information(QtGui.QApplication.activeWindow(),
                                                 u"TECHBOOK", msg, flags)
        assembly_recursion = response == QtGui.QMessageBox.Yes

        dialog = QtGui.QFileDialog(
            QtGui.QApplication.activeWindow(),
            "Select file to save your technical book"
        )
        # set option "DontUseNativeDialog"=True, as native Filedialog shows
        # misbehavior on Ubuntu 18.04 LTS. It works case sensitively, what is not wanted...
        if a2plib.getNativeFileManagerUsage():
            dialog.setOption(QtGui.QFileDialog.DontUseNativeDialog, False)
        else:
            dialog.setOption(QtGui.QFileDialog.DontUseNativeDialog, True)
        dialog.setDirectory(working_dir)
        dialog.setAcceptMode(QtGui.QFileDialog.AcceptSave)
        dialog.setNameFilter("Supported Formats (*.pdf);;All files (*.*)")
        if dialog.exec_():
            if a2plib.PYVERSION < 3:
                filename = unicode(dialog.selectedFiles()[0])
            else:
                filename = str(dialog.selectedFiles()[0])
            if not filename.endswith(".pdf"):
                filename += ".pdf"
        else:
            QtGui.QMessageBox.information(QtGui.QApplication.activeWindow(),
                                          u"Technical book generation aborted!",
                                          u"You have to give a file to save the book."
                                          )
            filename = None
        return assembly_recursion, filename

    def _getBookParameters(self, doc, standard_fields, editable_fields):
        """
        Look for #TECHINFO# spreadsheet, extract data from it and return them
        :param doc: Master document of the assembly.
        :param standard_fields: The function will put name of the standard
        fields (date, scale and sheet) in this dictionary.
        :param editable_fields: The function will put the editable fields
        in this dictionary.
        :return:
        The dictionaries standard_fields and editable_fields will be filled
        with data from the spreadsheet (empty before).
        """
        # We empty the returned values
        standard_fields.clear()
        editable_fields.clear()
        # We set the default standard values
        standard_fields['Stating_Page'] = 1
        standard_fields['Nb_Page_After'] = 0
        standard_fields['Date_Field'] = "FC-DATE"
        standard_fields['Scale_Field'] = "FC-SC"
        standard_fields['Sheet_Field'] = "FC-SH"
        # We get information from the spreadsheet
        sheet = doc.findObjects('Spreadsheet::Sheet',
                                Label=A2PCreateTechBook.INFO_SPREADSHEET_LABEL)
        if not sheet:
            return
        temp_dict = {}
        idx = 1
        run = True
        while run:
            try:
                key = sheet[0].get('A{}'.format(idx))
                val = sheet[0].get('B{}'.format(idx))
                temp_dict[key] = val
                idx += 1
            except ValueError:
                run = False
        # We set the ret_val
        for k, v in temp_dict.items():
            if k in A2PCreateTechBook.STANDARDS_FIELDS_NAME:
                if k == 'Stating_Page' or k == 'Nb_Page_After':
                    try:
                        v = int(v)
                    except ValueError:
                        v = standard_fields[k]
                standard_fields[k] = v
            else:
                editable_fields[k] = v

    def _getDocumentTechDraw(self, full_filename, treated):
        """
        Extract all the TechDraw elements from a document.
        :param full_filename: Filename with the document
        :param treated: list of document already treated
        :return: A tuple with the open document and a list of TechDraw
        the return value is (None, None) if the document does not exist
        if the document has no tech draw elements, the ret_val is (doc, [])
        """
        try:
            doc = FreeCAD.openDocument(full_filename)
        except OSError:
            print("File {} could not be opened !".format(full_filename))
            return []
        if doc in treated:
            return []
        treated.append(doc)
        td = doc.findObjects("TechDraw::DrawPage")
        if td:
            ret_val = [(doc, i) for i in td]
        else:
            ret_val = []
        return ret_val

    def _createTechDrawDocumentList(self, file_name, file_path,
                                    treated, recursive=True):
        """
        Create a list of document containing tech draw objects

        :param filename: Full path of the file to treat
        :param path: directory of the file to treat
        :param treated: list of document already treated
        :param recursive: do we work recursively

        :return: a list of tuple : (document, DrawPage)
        """
        #print("_createTechDrawDocumentList({}, {}, ...)".format(file_name, file_path))
        full_filename = str(Path(file_path) / Path(file_name))
        # We insert tech draw of the object given in parameter
        ret_val = self._getDocumentTechDraw(full_filename, treated)
        # We run throught the internal doc of the one given in parameter
        reader = FCdocumentReader()
        reader.openDocument(full_filename)
        for ob in reader.getA2pObjects():
            new_full_file = a2plib.findSourceFileInProject(ob.getA2pSource(), file_path)
            new_full_path, new_file_name = os.path.split(new_full_file)
            # if we are recursive, subassemblies are not treated here but in recursion
            # otherwise, we only include tech draw from the object himself
            if recursive and ob.isSubassembly():
                ret_val += self._createTechDrawDocumentList(new_file_name, new_full_path, treated, recursive)
            else:
                ret_val += self._getDocumentTechDraw(new_full_file, treated)
        return ret_val

    def _computeEditableFields(self, doc, page, templates_data,
                               date_field, scale_field):
        """
        Set the parameters in page template and recompute the doc
        :param doc: Document the draw page is from
        :param page: Draw page
        :param templates_data: all editable data
        :param date_field: name of the date field
        :param scale_field: name of the scale field
        :return:
        """
        texts = page.Template.EditableTexts

        # Python datetime does not support time ending with Z
        if doc.LastModifiedDate[-1] == 'Z':
            modified = datetime.datetime.fromisoformat(doc.LastModifiedDate[:-1])
        else:
            modified = datetime.datetime.fromisoformat(doc.LastModifiedDate)
        texts[date_field] = modified.strftime("%d/%m/%Y")
        texts[scale_field] = str(page.Scale)
        for k, v in templates_data.items():
            if k in texts:
                texts[k] = v
        page.Template.EditableTexts = texts
        page.recompute()

    def _computeTechDraw(self, doc_list, standard_fields, editable_fields):
        """
        Get all parameters from parameter dictionaries, set them to the TechDraw
        and compute the doc.
        :param doc_list: list of tuple (doc, techDrawPage)
        :param standard_fields: Standard parameters
        :param editable_fields: editable parameters set by user
        :return:
        """
        # We calculate the number of page
        nb_page = standard_fields['Stating_Page'] + len(doc_list) + standard_fields['Nb_Page_After'] - 1
        # We set standard fields
        date_field = standard_fields['Date_Field']
        scale_field = standard_fields['Scale_Field']
        sheet_field = standard_fields['Sheet_Field']
        templates_data = editable_fields.copy()
        page_edited = standard_fields['Stating_Page']
        for doc, page in doc_list:
            templates_data[sheet_field] = "{} / {}".format(page_edited, nb_page)
            self._computeEditableFields(doc, page, templates_data, date_field, scale_field)
            page_edited += 1

    def _createPDFFile(self, base_doc, doc_list, filename):
        """
        Create all the PDF files from the tech draw and merge them to a uniq file.
        :param base_doc: Master A2P file
        :param doc_list: list of tuple (doc, techDrawPage)
        :param filename: file to save the final pdf
        :return: return the list of files created
        """
        nb = 1
        pdf_merger = PdfFileMerger()
        for doc, page in doc_list:
            pdf_file = base_doc.getTempFileName("Page{}".format(nb)) + ".pdf"
            if not page.Visibility:
                page.ViewObject.doubleClicked()
            TechDrawGui.exportPageAsPdf(page, pdf_file)
            pdf_merger.append(pdf_file)
            nb += 1
        with open(filename, "wb") as output_file:
            pdf_merger.write(output_file)

    def Activated(self):
        doc = FreeCAD.activeDocument()
        if not doc:
            QtGui.QMessageBox.information(QtGui.QApplication.activeWindow(),
                                          u"No active document found!",
                                          u"You have to open a FCStd file first."
                                          )
            return
        complete_file_path = doc.FileName
        path, doc_filename = os.path.split(complete_file_path)

        recursion, filename = self._getUserParameters(path)
        if not filename:
            return
        standard_fields = {}
        editable_fields = {}
        self._getBookParameters(doc, standard_fields, editable_fields)
        doc_list = self._createTechDrawDocumentList(doc_filename, path,
                                                    treated=[],
                                                    recursive=recursion)
        self._computeTechDraw(doc_list, standard_fields, editable_fields)
        self._createPDFFile(doc, doc_list, filename)
        QtGui.QMessageBox.information(QtGui.QApplication.activeWindow(),
                                      u"TechBook completed",
                                      u"File {} created.".format(filename)
                                      )

    def GetResources(self):
        return {
            'Pixmap': ':/icons/a2p_TechBook.svg',
            'MenuText': QT_TRANSLATE_NOOP(
                "A2plus_CreateTechBook",
                "Create a file containing all the Tech Draws in the assembly."),
            'ToolTip': toolTip
            }


FreeCADGui.addCommand('a2p_createTechBook', A2PCreateTechBook())
