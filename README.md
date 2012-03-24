VisioDrawing2SplittedPDFs
=========================

Exports each sheet of a .vsd as separate PDFs using Apache PDFBox.
The directory used is the directory of the .vsd.

Tested with Microsoft Visio 2007.

Background
----------
I have a Microsoft Visio drawing with multiple pages.
In LaTeX, I cannot access each page indivudally, but have to use PDF
as intermediate format. Therefore, I save the PDF and extract the
necessary pages and rename each to the name of the respective sheet.
This macro automates that using Apache PDFBox.

Requirements
------------
 * Java
 * Apache PDFBox 1.6.0 app (http://pdfbox.apache.org/download.html -> pdfbox-app-1.6.0.jar)

Installation
------------
 * move pdfbox-app-1.6.0.jar to c:\apache

Usage
-----
Visio does not offer a normal.dot. Therefore, it is not possible, to add the 
macro permanently to the toolbar (using VBA). Each document has to contain the
code for the splittedPDFexport:

 * Press ALT+F11 to open the VBA editor
 * Press CTRL+M and import splittedPDFexport.bas
 * Navigate to "ThisDocument (<document name>)"
 * Add following code
 
        Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
          Call splittedPDFexport.SetToolbar
        End Sub 
    
 * To get rid of macro warnings, sign the VBA macro of splittedPDFexport
   * Follow the guideline by microsoft (http://msdn.microsoft.com/en-us/library/aa163622%28office.10%29.aspx)

Customization
-------------
 * The toolbar to put the button is configured by the constant toolbarCaption
 * The caption of the button is configured by the constant buttonCaption
 * The position of the button in the toolbar can only be changed in the code.
   Look for "Set vsoToolbarItem = vsoToolbarItems.Add" and change "Add" to ".AddAt(<position>)"
 * The path to the PDFBox application is configured by the constant pdfboxappCMD
   
Notes on the implementation
---------------------------
 * The macro could also be stored in a .vss file. I was not possible, however,
   to call the makro from the main document
   
License
-------
VisioDrawing2SplittedPDFs
  copyright 2012 Oliver Kopp. All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

   1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
   2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
   3. The name of the author may not be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
