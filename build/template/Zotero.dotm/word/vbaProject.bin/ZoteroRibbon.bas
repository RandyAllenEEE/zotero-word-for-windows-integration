Attribute VB_Name = "ZoteroRibbon"
' ***** BEGIN LICENSE BLOCK *****
'
' Copyright (c) 2015  Zotero
'                     Center for History and New Media
'                     George Mason University, Fairfax, Virginia, USA
'                     http://zotero.org
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' ***** END LICENSE BLOCK *****

Option Explicit

Sub ZoteroRibbonAddEditCitation(button As IRibbonControl)
    Call ZoteroAddEditCitation
End Sub

Sub ZoteroRibbonAddNote(button As IRibbonControl)
    Call ZoteroAddNote
End Sub

Sub ZoteroRibbonAddEditBibliography(button As IRibbonControl)
    Call ZoteroAddEditBibliography
End Sub

Sub ZoteroRibbonSetDocPrefs(button As IRibbonControl)
    Call ZoteroSetDocPrefs
End Sub

Sub ZoteroRibbonRefresh(button As IRibbonControl)
    Call ZoteroRefresh
End Sub

Sub ZoteroRibbonRemoveCodes(button As IRibbonControl)
    Call ZoteroRemoveCodes
End Sub

Sub ZoteroRibbonConvert(button As IRibbonControl)
    ConvertTextCitations
End Sub

Private Sub ConvertTextCitations()
    Dim regEx As Object
    Dim matches As Object
    Dim match As Object
    Dim docContent As String
    Dim i As Integer
    Dim rng As Range
    Dim fld As Field
    Dim itemKey As String
    Dim json As String
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    ' Pattern matches: [1](zotero://select/items/KEY)
    ' Default pattern provided allows for standard Markdown links with Zotero select URIs
    Dim defaultPattern As String
    defaultPattern = "\[.*?\]\(zotero://.*?/items/([A-Z0-9]{8})\)"
    
    Dim userPattern As String
    userPattern = InputBox("Enter Regex Pattern (Group 1 must be the Item Key):", "Convert Text to Zotero Citations", defaultPattern)
    
    If userPattern = "" Then Exit Sub
    
    regEx.Pattern = userPattern
    
    docContent = ActiveDocument.Content.Text
    
    If Not regEx.Test(docContent) Then
        MsgBox "No text citations found.", vbInformation, "Zotero Converter"
        Exit Sub
    End If
    
    Set matches = regEx.Execute(docContent)
    
    ' Process backwards to preserve range indices
    For i = matches.Count - 1 To 0 Step -1
        Set match = matches.Item(i)
        itemKey = match.SubMatches(0)
        
        ' VBScript RegExp returns 0-based index. Word Range is also character-based.
        ' Note: Word may count specific hidden chars differently, but usually this aligns for body text.
        Set rng = ActiveDocument.Range(match.FirstIndex, match.FirstIndex + match.Length)
        
        ' Create Minimal CSL JSON
        ' We escape quotes by doubling them in VBA strings
        json = "{""citationID"":""CIT_" & itemKey & "_" & match.FirstIndex & """,""properties"":{""formattedCitation"":""[Loading...]"",""plainCitation"":""[Loading...]""},""citationItems"":[{""id"":""" & itemKey & """,""uris"":[""http://zotero.org/users/local/0/items/" & itemKey & """],""itemData"":{""id"":""" & itemKey & """,""type"":""article-journal"",""title"":""Loading...""}}],""schema"":""https://github.com/citation-style-language/schema/raw/master/schemas/input/csl-data.json""}"
        
        ' Attempt to add field. Fields.Add replaces the range automatically.
        ' STRATEGY: Zotero C++ code uses wdFieldQuote (35) first to prevent range issues, 
        ' then changes the Code to " ADDIN ... ". We mimic this behavior.
        
        On Error Resume Next
        ' 1. Create Quote field
        Set fld = ActiveDocument.Fields.Add(rng, 35, "TEMP", False)
        
        If Not fld Is Nothing Then
            ' 2. Change to ADDIN by setting Code.Text
            fld.Code.Text = " ADDIN ZOTERO_ITEM CSL_CITATION " & json & " "
            
            ' 3. Set Result
            fld.Result.Text = "[Loading...]"
        End If
        
        If Err.Number <> 0 Then
            Debug.Print "Failed to convert at index " & match.FirstIndex & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    
    MsgBox "Converted " & matches.Count & " citations." & vbCrLf & _
           "Please click 'Refresh' to update them.", vbInformation, "Zotero Converter"
End Sub

Sub ZoteroTabLabel(tb As IRibbonControl, ByRef returnedVal)
    Dim ver As Double
    ver = Val(Application.Version)
    If ver >= 15 And ver < 16 Then
        returnedVal = "ZOTERO"
    Else
        returnedVal = "Zotero"
    End If
End Sub
