Attribute VB_Name = "modFileFunctions"
Option Explicit
Global gTreeKey As Long
Public Function CountChar(StringX As String, CharX As String) As Long
'Count the occurance of a character in a string

    Dim ii As Long
    
    For ii = 1 To Len(StringX)
        If Mid(StringX, ii, 1) = CharX Then CountChar = CountChar + 1
    Next ii
    
    'CountChar = CountChar + 1 ' For my needs , it can never be 0
    
End Function
Public Function GetXMLFileOpen(DlgX As Object) As String
'Prompts the CommonDialog for a XML File to Open
Dim sTargetFile As String
Dim sSourceFile As String
Dim Mycheck


DlgX.Filter = "XML Files(*.xml)|*.xml|All Files(*.*)|*.*"
DlgX.DefaultExt = ".xml"
DlgX.ShowOpen
sTargetFile = DlgX.FileName

If sTargetFile = "" Then Exit Function

Mycheck = Dir(sTargetFile)


GetXMLFileOpen = sTargetFile

End Function
Public Function OpenString(sFileName As String) As String
'Opens a file and reads it into a string
Dim Printline As String
Dim MyString As String

Open sFileName For Input As #1    ' Open file for output.
    Do While Not EOF(1) ' Loop until end of file.
    Input #1, MyString    ' Read data into two variables.
    Printline = Printline & MyString ' Print data to Debug window.
    Loop
Close #1 ' Close the newly created file

OpenString = Printline

End Function


Public Function GetXMLFileSave() As String
Dim sTargetFile As String
Dim sSourceFile As String
Dim Mycheck


frmMDIMain.dlg.Filter = "XML Files(*.xml)|*.xml"
frmMDIMain.dlg.DefaultExt = ".xml"
frmMDIMain.dlg.ShowSave
sTargetFile = frmMDIMain.dlg.FileName

If sTargetFile = "" Then Exit Function

Mycheck = Dir(sTargetFile)

If Mycheck = frmMDIMain.dlg.FileTitle Then
    Dim result
    result = MsgBox("This file already exists. Overrite ?", vbYesNo, "File Exists")
    If result = vbYes Then
        Kill sTargetFile
        DoEvents
    Else
        Exit Function
    End If
End If

GetXMLFileSave = sTargetFile

End Function
Public Function GetKey() As String
    gTreeKey = gTreeKey + 1
    GetKey = "Key" & gTreeKey
End Function

Public Sub ColourXML(StringX)
    Dim lngLastPos As Long
    Dim lngLength As Long
    
With frmXML

    .rtfMain.Text = ""
    .rtfMain.SelColor = vbBlack
    .rtfMain.SelBold = True
    
    '.rtfMain.LoadFile strXMLLocation, 1
    .rtfMain.Text = StringX
   
   .rtfMain.span ("<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>")
   lngLastPos = .rtfMain.SelLength - 1
   .rtfMain.SelColor = vbBlue
   .rtfMain.SelBold = False
   
    Do While lngLastPos > -1
        lngLastPos = .rtfMain.Find("<", lngLastPos + 1)
        If lngLastPos = -1 Then Exit Do
        .rtfMain.SelColor = vbBlue
        .rtfMain.SelBold = False
        
        lngLength = (.rtfMain.Find(">", lngLastPos)) - lngLastPos
        .rtfMain.SelStart = lngLastPos + 1
        .rtfMain.SelLength = lngLength
        .rtfMain.SelColor = &HC0&
        .rtfMain.SelBold = False
    Loop
    
    lngLastPos = 0
    
    Do While lngLastPos > -1
        lngLastPos = .rtfMain.Find(">", lngLastPos + 1)
        .rtfMain.SelColor = vbBlue
        .rtfMain.SelBold = False
    Loop
    
    lngLastPos = 0
    
    Do While lngLastPos > -1
        lngLastPos = .rtfMain.Find("</", lngLastPos + 1)
        .rtfMain.SelColor = vbBlue
        .rtfMain.SelBold = False
    Loop
    
    lngLastPos = 0
    
    Do While lngLastPos > -1
        lngLastPos = .rtfMain.Find("=", lngLastPos + 1)
        .rtfMain.SelColor = vbBlue
        .rtfMain.SelBold = False
    Loop
    
    lngLastPos = 0
    
    Do While lngLastPos > -1
        lngLastPos = .rtfMain.Find("'", lngLastPos + 1)
        .rtfMain.SelColor = vbBlue
        .rtfMain.SelBold = False
        
        If lngLastPos = -1 Then Exit Do
        lngLength = lngLastPos
        lngLastPos = .rtfMain.Find("'", lngLastPos + 1)
        
        .rtfMain.SelStart = lngLength + 1
        .rtfMain.SelLength = (lngLastPos - lngLength) - 1
        .rtfMain.SelColor = &H8000&
        .rtfMain.SelBold = True
    Loop
    
    lngLastPos = 0
    
    Do While lngLastPos > -1
        lngLastPos = .rtfMain.Find("'", lngLastPos + 1)
        .rtfMain.SelColor = vbBlue
        .rtfMain.SelBold = False
    Loop
    
    .rtfMain.SelStart = 0
    '.rtfMain.SetFocus
End With
End Sub

