Attribute VB_Name = "splittedPDFexport"
'****************************************************************************
'License:
'
'VisioDrawing2SplittedPDFs
'  copyright 2012 Oliver Kopp. All rights reserved.
'
'Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
'
'   1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
'   2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
'   3. The name of the author may not be used to endorse or promote products derived from this software without specific prior written permission.
'
'THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'****************************************************************************

'Version 0.1 - 2012-03-24

'Changelog
'
'Version 0.1 - 2012-03-24
' * first public release

Option Explicit

'****************************************************************************
'Configuration
'****************************************************************************

Private Const pdfboxappCMD = "java -jar c:\apache\pdfbox-app-1.6.0.jar"

'Toolbar to add button to
'Private Const toolbarCaption = "Splitted PDF Export" 'If this option is used, the toolbar floats around and the user has to dock it by hand
Private Const toolbarCaption = "Standard"

Private Const buttonCaption = "PDF Export"

'****************************************************************************
'Implementation
'****************************************************************************

'Macro for executing the PDF export
Public Sub PDFexport()
    'save whole file as .pdf
    'Application.ActiveDocument.ExportAsFixedFormat visFixedFormatPDF, pdfName, visDocExIntentPrint, visPrintAll ', 1, 7, False, True, True, True, False
    
    'create .bat file for splitting into pdfs
    
    Dim intCounter As Integer
    Dim vsoPages As Visio.Pages
    
    Dim pdfName As String
    Dim fileNameWithoutExtension As String
    Dim filePath As String
    pdfName = ActiveDocument.FullName
    Dim i, j As Long
    i = InStrRev(pdfName, ".")
    j = InStrRev(pdfName, "\")
    filePath = Left(pdfName, j)
    fileNameWithoutExtension = Mid(pdfName, j + 1, i - j - 1)
    pdfName = Left(pdfName, i) + "pdf"
    
    'Get the Pages collection for the active document.
    Set vsoPages = ActiveDocument.Pages

    Dim fnum As Integer
    fnum = FreeFile
    Dim batFile As String
    batFile = filePath + "runme.bat"
    Open batFile For Output As fnum

    Dim splitNameInQuotes As String
    Dim finalNameInQuotes As String
    
    'change to target dir
    Print #fnum, "cd " + filePath
    Dim drive As String
    drive = Left(filePath, 2)
    Print #fnum, drive
    
    'split pdf
    Print #fnum, pdfboxappCMD + " PDFSplit " + pdfName

    'rename splitted pdf to right name
    For intCounter = 1 To vsoPages.Count
        splitNameInQuotes = """" + fileNameWithoutExtension + "-" + Format(intCounter - 1) + ".pdf"""
        finalNameInQuotes = """" + vsoPages.Item(intCounter).Name + ".pdf"""
        Print #fnum, "del " + finalNameInQuotes
        Print #fnum, "ren " + splitNameInQuotes + " " + finalNameInQuotes
    Next intCounter
    Close fnum
    
    Dim retVal
    retVal = Shell(batFile, vbNormalNoFocus)
End Sub

'Add button to toolbar
'Adapted from http://www.office-loesung.de/ftopic303582_0_0_asc.php
Public Sub SetToolbar()
    Dim vsoUIObject As Visio.UIObject
    Dim vsoToolbarSet As Visio.ToolbarSet
    Dim vsoToolbar As Visio.Toolbar
    Dim vsoToolbarItems As Visio.ToolbarItems
    Dim vsoToolbarItem As Visio.ToolbarItem
    Dim lAnz As Long, lPos As Long
 
    'Get the UIObject object --> mit .BuiltInToolbars werden die
    'eingebauten Toolbars zurückgeliefert, siehe VBA-Hilfe
    Set vsoUIObject = Visio.Application.CustomToolbars     'Get the drawing window toolbar sets.
 
    'NOTE: Use ItemAtID to get the toolbar set.
    'Using vsoUIObject.ToolbarSets(visUIObjSetDrawing) will not work.
    Set vsoToolbarSet = vsoUIObject.ToolbarSets.ItemAtID(visUIObjSetDrawing)
    
    'Find the toolbar
    'If not found: add as new
    '
    'Get the ToolbarItems collection.
    lAnz = vsoToolbarSet.Toolbars.Count - 1
    'Richtige Toolbar suchen
    '--> Text "MyToolbar" mit dem richtigen Namen ersetzen
    Dim found As Boolean
    found = False
    For lPos = 0 To lAnz
        If vsoToolbarSet.Toolbars(lPos).Caption = toolbarCaption Then
            lAnz = lPos
            found = True
        End If
    Next lPos
    If found Then
        Set vsoToolbar = vsoToolbarSet.Toolbars(lAnz)
    Else
        Set vsoToolbar = vsoToolbarSet.Toolbars.Add
        With vsoToolbar
            .Caption = toolbarCaption
            .Visible = True
        End With
    End If
        
    Set vsoToolbarItems = vsoToolbar.ToolbarItems
    '' Um sicher sauber zu sein, zuerst alle Knöpfe löschen
    'For lPos = vsoToolbarItems.Count - 1 To 0 Step -1
    '    vsoToolbarItems(lPos).Delete
    'Next lPos
    Set vsoToolbarItem = vsoToolbarItems.Add
    With vsoToolbarItem
        .CntrlType = visCtrlTypeBUTTON
        .Style = visButtonCaption
        .Caption = buttonCaption
        .AddOnName = "splittedPDFexport.PDFexport"
        .AddOnArgs = ""
    End With
    ThisDocument.SetCustomToolbars vsoUIObject
End Sub
