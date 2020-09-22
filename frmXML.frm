VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmXML 
   Caption         =   "XML Wizards"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8025
   Icon            =   "frmXML.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlg 
      Left            =   1575
      Top             =   5775
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlXML 
      Left            =   2100
      Top             =   5670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":0626
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":0942
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":0C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":0F7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":1296
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmXML.frx":15B2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tbsMain 
      Height          =   6315
      Left            =   2835
      TabIndex        =   1
      Top             =   0
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   11139
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "View"
      TabPicture(0)   =   "frmXML.frx":18CE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdXML"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Print Preview"
      TabPicture(1)   =   "frmXML.frx":18EA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPrint"
      Tab(1).Control(1)=   "cmdGenAll"
      Tab(1).Control(2)=   "cmdGenVisible"
      Tab(1).Control(3)=   "Browser"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Search"
      TabPicture(2)   =   "frmXML.frx":1906
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraSearch"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraSearch 
         Caption         =   "Search the Tree"
         Height          =   1485
         Left            =   -74895
         TabIndex        =   7
         Top             =   525
         Width           =   4950
         Begin VB.TextBox txtSearch 
            Height          =   330
            Left            =   1995
            TabIndex        =   9
            Top             =   420
            Width           =   2850
         End
         Begin VB.CommandButton cmdTreeSearch 
            Caption         =   "Find It"
            Height          =   435
            Left            =   3675
            TabIndex        =   8
            Top             =   840
            Width           =   1170
         End
         Begin VB.Label lblSearch 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Value to find"
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   105
            TabIndex        =   10
            Top             =   420
            Width           =   1905
         End
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   435
         Left            =   -73320
         TabIndex        =   6
         Top             =   5775
         Width           =   1695
      End
      Begin VB.CommandButton cmdGenAll 
         Caption         =   "Get all Nodes"
         Height          =   435
         Left            =   -71535
         TabIndex        =   5
         Top             =   5775
         Width           =   1485
      End
      Begin VB.CommandButton cmdGenVisible 
         Caption         =   "Get visible Nodes"
         Height          =   435
         Left            =   -74895
         TabIndex        =   4
         Top             =   5775
         Width           =   1485
      End
      Begin MSFlexGridLib.MSFlexGrid grdXML 
         Height          =   5580
         Left            =   105
         TabIndex        =   3
         Top             =   525
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   9843
         _Version        =   393216
         FixedRows       =   0
         BackColorBkg    =   12632256
         BorderStyle     =   0
      End
      Begin SHDocVwCtl.WebBrowser Browser 
         Height          =   5265
         Left            =   -74895
         TabIndex        =   2
         Top             =   420
         Width           =   4950
         ExtentX         =   8731
         ExtentY         =   9287
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin MSComctlLib.TreeView trvXML 
      Height          =   5580
      Left            =   105
      TabIndex        =   0
      Top             =   525
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   9843
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlXML"
      Appearance      =   1
   End
   Begin VB.Image imgLogo 
      Height          =   750
      Left            =   -105
      Picture         =   "frmXML.frx":1922
      Top             =   -105
      Width           =   3000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open XML"
      End
      Begin VB.Menu xx 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintTree 
         Caption         =   "Print Visible Tree"
      End
   End
End
Attribute VB_Name = "frmXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFileName As String
Dim sXMLstring As String
Dim newXMLstring As String
Dim gbCheck As Boolean
Dim objXML As New DOMDocument
Dim domNode As IXMLDOMNode
Dim objChange As Object
Dim ParentName As String
Dim ii As Integer
Dim bb As Integer
Dim iLastCol As Integer
Dim iLastRow As Integer

Private Sub cmdGenerateXML_Click()
txtXML = objXML.xml
End Sub


Private Sub cmdGenAll_Click()
Call PrintTree(True)
End Sub

Private Sub cmdGenVisible_Click()
Call PrintTree
End Sub

Private Sub cmdPrint_Click()
Browser.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
DoEvents
End Sub

Private Sub cmdTreeSearch_Click()
'Loop through the tree to find a value
txtSearch = Trim(LCase(txtSearch))
For ii = 1 To trvXML.Nodes.Count
    If LCase(trvXML.Nodes(ii).Text) Like "%" & txtSearch & "%" Or _
        LCase(trvXML.Nodes(ii).Text) Like txtSearch & "%" Or _
        LCase(trvXML.Nodes(ii).Text) Like "%" & txtSearch Or _
        LCase(trvXML.Nodes(ii).Text) = txtSearch Then
        trvXML.Nodes(ii).Image = 7
        trvXML.Nodes(ii).EnsureVisible
    End If
Next ii

End Sub

Private Sub Form_Load()
Me.Top = Screen.Height / 2 - Me.Height / 2
Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

Private Sub grdXML_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
iLastCol = grdXML.MouseCol
iLastRow = grdXML.MouseRow
End Sub

Private Sub mnuOpen_Click()
    sFileName = GetXMLFileOpen(dlg)
    sXMLstring = OpenString(sFileName)
    ParentName = dlg.FileTitle
    Dim NodX As Node
    trvXML.Nodes.Clear
    Set NodX = trvXML.Nodes.Add(, , dlg.FileTitle, dlg.FileTitle, 1) ' Parent
    
gbCheck = objXML.loadXML(sXMLstring)

If gbCheck = False Then
    MsgBox "Unable to load " & dlg.FileTitle & " as a valid XML string"
    Exit Sub
End If

gTreeKey = 0

For ii = 0 To objXML.childNodes.length - 1
    Set NodX = trvXML.Nodes.Add(ParentName, tvwChild, ParentName & ii, objXML.childNodes.Item(ii).nodeName, 2) ' Child
    Dim objXMLchild As New objTreeBuild
    gbCheck = objXMLchild.BuildTree(objXML.childNodes.Item(ii), trvXML, ParentName & ii, objXML)
    Set objXMLchild = Nothing
Next ii

'Call ColourXML(objXML.xml)
'Browser.navigate "file://" & dlg.FileName
'Set NodX = trvXML.Nodes.Add(ObjX.ADOcon.DefaultDatabase, tvwChild, "Tables", "Tables", "tables") ' Parent
Call BuildGrid
    
End Sub


Private Sub mnuPrintTree_Click()
Call PrintTree
End Sub

Private Sub trvXML_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim StringX As String

If trvXML.SelectedItem.Image = imgXML.ListImages(4).Picture Then
    gbCheck = ChangeAttributeValue(NewString, trvXML.SelectedItem.Parent.Key, trvXML.SelectedItem.Parent.Text)
Else
    gbCheck = ChangeElementValue(NewString, trvXML.SelectedItem.Parent.Key, trvXML.SelectedItem.Parent.Text)
End If

'txtXML = objXML.xml
End Sub

Private Function ChangeElementValue(sValue As String, sKey As String, sParent As String) As Boolean

'Change the value of an Element
Set objChange = objXML.selectNodes("//" & sParent)
For ii = 0 To objChange.length - 1
    For bb = 0 To objChange(ii).Attributes.length - 1
        If objChange(ii).Attributes(bb).Text = sKey Then
            objChange(ii).Text = sValue
        End If
    Next bb
Next ii
End Function
Private Function ChangeAttributeValue(sValue As String, sKey As String, sParent As String) As Boolean

'Change the value of an Attributes
Set objChange = objXML.selectNodes("//" & sParent)
For ii = 0 To objChange.length - 1
    For bb = 0 To objChange(ii).Attributes.length - 1
        If objChange(ii).Attributes(bb).Text = sKey Then
            objChange(ii).Attributes(bb).Text = sValue
        End If
    Next bb
Next ii
End Function
Private Sub PrintTree(Optional bWhole As Boolean)
Dim sPath As String
On Error Resume Next
Dim iRowCount As Long
Dim iDepth As Integer
Dim HTML As String
Dim bb As Integer
Dim sImage As String

iRowCount = -2
grdXML.Clear
grdXML.cols = 3
grdXML.rows = trvXML.GetVisibleCount
grdXML.ColWidth(0) = 250
grdXML.ColWidth(1) = 3000
grdXML.ColWidth(2) = 500
HTML = HTML & "<BODY bgcolor='#c0c0c0'>"
HTML = HTML & "<img src='logo.GIF'><BR>"
'Browser.document.body.innerHTML = "Hallo"
For ii = 0 To trvXML.Nodes.Count
    If trvXML.Nodes(ii).Visible = True Or bWhole = True Then
        iRowCount = iRowCount + 1
        grdXML.Row = iRowCount
        grdXML.Col = 1
        
        iDepth = CountChar(trvXML.Nodes(ii).FullPath, "\")
        For bb = 1 To iDepth
            HTML = HTML & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        Next bb
   
        
        Select Case trvXML.Nodes(ii).Image
            Case 1
                sImage = "FileRoot.ico"
            Case 2
                sImage = "Record.ico"
            Case 3
                sImage = "Element.ico"
            Case 4
                sImage = "Property.ico"
            Case 5
                sImage = "PropText.ico"
            Case 6
                sImage = "EleText.ico"
        End Select
        
        HTML = HTML & "<img src='" & sImage & "' height=16 width=16 vallign='top'> " & trvXML.Nodes(ii).Text
        HTML = HTML & "<BR>"
        
    End If
Next ii


Kill "c:\testhtml\test.html"

'Now Rewrite the whole file
Open App.Path & "\XML_123_Leave.html" For Output As #2 ' Open file for output."
    'For ii = 0 To 10
        Print #2, HTML  'Save Record
    'Next ii
Close #2    ' Close file.

'rtfPrint.SaveFile "c:\testhtml\test.html"
For ii = 1 To 10000
DoEvents
Next ii

'Browse.navigate "file://c:\testhtml\test.html"
'Me.Refresh

Browser.navigate "file://" & App.Path & "\XML_123_Leave.html" 'Just a dummy to get the DOM initialised !!! - Clever eh ?
DoEvents
'Browser.document.body.innerHTML = HTML
DoEvents
Browser.Refresh
DoEvents
End Sub
Private Sub BuildGrid()
On Error Resume Next
Dim iRowCount As Long

iRowCount = -2
grdXML.Clear
grdXML.cols = 3
grdXML.rows = trvXML.GetVisibleCount
grdXML.ColWidth(0) = 250
grdXML.ColWidth(1) = 3000
grdXML.ColWidth(2) = 500

For ii = 0 To trvXML.Nodes.Count
    If trvXML.Nodes(ii).Visible = True Then
        iRowCount = iRowCount + 1
        grdXML.Row = iRowCount
        grdXML.Col = 1
        
        If trvXML.Nodes(ii).Image = 5 Or trvXML.Nodes(ii).Image = 6 Then
            grdXML.CellBackColor = vbWhite
        Else
            grdXML.CellBackColor = grdXML.BackColorBkg
        End If
        
'        Set grdXML.CellPicture = imgXML.ListImages(trvXML.Nodes(ii).Image).Picture
'        grdXML.CellPictureAlignment = flexAlignCenterCenter
        grdXML.TextMatrix(iRowCount, 1) = trvXML.Nodes(ii).Text
        grdXML.TextMatrix(iRowCount, 2) = ii
    End If
Next ii


End Sub

Private Sub trvXML_Collapse(ByVal Node As MSComctlLib.Node)
Call BuildGrid
End Sub

Private Sub trvXML_Expand(ByVal Node As MSComctlLib.Node)
Call BuildGrid
End Sub

Private Sub trvXML_KeyDown(KeyCode As Integer, Shift As Integer)
'Call BuildGrid
End Sub

Private Sub trvXML_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call BuildGrid
End Sub
'Private Sub PrintTree(Optional bWhole As Boolean)
'Dim sPath As String
'On Error Resume Next
'Dim iRowCount As Long
'Dim iDepth As Integer
'Dim HTML As String
'Dim bb As Integer
'Dim sImage As String
'
'iRowCount = -2
'grdXML.Clear
'grdXML.cols = 3
'grdXML.rows = trvXML.GetVisibleCount
'grdXML.ColWidth(0) = 250
'grdXML.ColWidth(1) = 3000
'grdXML.ColWidth(2) = 500
'HTML = HTML & "<BODY bgcolor='#c0c0c0'>"
'HTML = HTML & "<img src='logo.GIF'><BR><TABLE BORDER=1>"
''Browser.document.body.innerHTML = "Hallo"
'For ii = 0 To trvXML.Nodes.Count
'    If trvXML.Nodes(ii).Visible = True Then
'        HTML = HTML & "<TR>"
'        iRowCount = iRowCount + 1
'        grdXML.Row = iRowCount
'        grdXML.Col = 1
'
'        iDepth = CountChar(trvXML.Nodes(ii).FullPath, "\")
'        For bb = 1 To iDepth
'            HTML = HTML & "<TD>&nbsp;</TD>"
'        Next bb
'
'
'        Select Case trvXML.Nodes(ii).Image
'            Case 1
'                sImage = "FileRoot.ico"
'            Case 2
'                sImage = "Record.ico"
'            Case 3
'                sImage = "Element.ico"
'            Case 4
'                sImage = "Property.ico"
'            Case 5
'                sImage = "PropText.ico"
'            Case 6
'                sImage = "EleText.ico"
'        End Select
'
'        HTML = HTML & "<TD><img src='" & sImage & "' height=16 width=16 vallign='top'>" & trvXML.Nodes(ii).Text & "</TD>" '& "<BR>"
'        HTML = HTML & "</TR>"
'
'    End If
'Next ii
'HTML = HTML & "</TABLE>"
'
'Kill "c:\testhtml\test.html"
'
''Now Rewrite the whole file
'Open App.Path & "\XML_123_Leave.html" For Output As #2 ' Open file for output."
'    'For ii = 0 To 10
'        Print #2, HTML  'Save Record
'    'Next ii
'Close #2    ' Close file.
'
''rtfPrint.SaveFile "c:\testhtml\test.html"
'For ii = 1 To 10000
'DoEvents
'Next ii
'
''Browse.navigate "file://c:\testhtml\test.html"
''Me.Refresh
'
'Browser.navigate "file://" & App.Path & "\XML_123_Leave.html" 'Just a dummy to get the DOM initialised !!! - Clever eh ?
'DoEvents
''Browser.document.body.innerHTML = HTML
'DoEvents
'Browser.Refresh
'DoEvents
'End Sub
