VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarChildString As Variant 'local copy
Private mvarTag As Variant 'local copy
Private mvarAttributes As New Collection 'local copy
Private mvarElements As New Collection 'local copy
Private mvarWholeString As Variant 'local copy
Private mvarRootTag As Variant 'local copy
Private mvarTreeView As Object 'local copy
Private bCheck As Boolean
Private mvarTreeKey As Variant 'local copy
Private mvarTreeRoot As Variant 'local copy
Public Property Let TreeRoot(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TreeRoot = 5
    mvarTreeRoot = vData
End Property

Public Property Get TreeRoot() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TreeRoot
    If IsObject(mvarTreeRoot) Then
        Set TreeRoot = mvarTreeRoot
    Else
        TreeRoot = mvarTreeRoot
    End If
End Property



Public Property Let TreeKey(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TreeKey = 5
    mvarTreeKey = vData
End Property

Public Property Get TreeKey() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TreeKey
    If IsObject(mvarTreeKey) Then
        Set TreeKey = mvarTreeKey
    Else
        TreeKey = mvarTreeKey
    End If
End Property



Public Property Set TreeView(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.TreeView = Form1
    Set mvarTreeView = vData
End Property


Public Property Get TreeView() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TreeView
    Set TreeView = mvarTreeView
End Property



Public Property Let RootTag(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RootTag = 5
    mvarRootTag = vData
End Property


Public Property Get RootTag() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RootTag
    If IsObject(mvarRootTag) Then
        Set RootTag = mvarRootTag
    Else
        RootTag = mvarRootTag
    End If
End Property



Public Property Let WholeString(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WholeString = 5
    mvarWholeString = vData
End Property

Public Property Get WholeString() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WholeString
    If IsObject(mvarWholeString) Then
        Set WholeString = mvarWholeString
    Else
        WholeString = mvarWholeString
    End If
End Property



Public Property Set Elements(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Elements = Form1
    Set mvarElements = vData
End Property


Public Property Get Elements() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Elements
    Set Elements = mvarElements
End Property



Public Property Set Attributes(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Attributes = Form1
    Set mvarAttributes = vData
End Property


Public Property Get Attributes() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Attributes
    Set Attributes = mvarAttributes
End Property



Public Property Let Tag(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tag = 5
    mvarTag = vData
End Property

Public Property Get Tag() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    If IsObject(mvarTag) Then
        Set Tag = mvarTag
    Else
        Tag = mvarTag
    End If
End Property



Public Property Let ChildString(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ChildString = 5
    mvarChildString = vData
End Property


Public Property Get ChildString() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ChildString
    If IsObject(mvarChildString) Then
        Set ChildString = mvarChildString
    Else
        ChildString = mvarChildString
    End If
End Property



Private Function GetTagName(ByVal StringX As String, Optional lStartPos As Long) As String
'This function will get the root tag name of the string from a given position

Dim ii As Integer
If lStartPos < 2 Then lStartPos = 1
If lStartPos > Len(StringX) Then lStartPos = Len(StringX)

For ii = lStartPos To Len(StringX)
    If Mid(StringX, ii, 1) = "<" And (ii > lStartPos + 2) Then
        Exit Function
    Else
        If Mid(StringX, ii, 1) = " " Or Mid(StringX, ii, 1) = ">" Or Mid(StringX, ii, 1) = "/" Then
            GetTagName = Mid(StringX, 2, ii - 2)
            Exit Function
        End If
    End If
Next ii

End Function

Private Function CountRootTags(StringX As String) As Integer
'The will count the number of root tags for the string

Dim sTagString As String
Dim sDataString As String
Dim sTempString As String
Dim sEndTag As String
Dim sStartTag As String
Dim iPos As Long
Dim ii As Integer

While Len(sTempString) < Len(StringX)
    sTagString = GetTagName(StringX, 1)
    sStartTag = GetTagName(Right(StringX, Len(StringX) - iPos), 1)
    sEndTag = "</" & sStartTag & ">"
    'sDataString = GetInBetween(StringX, "<" & GetFullTag(mvarWholeString, 1) & ">", "</" & mvarTag & ">")
    sDataString = GetInBetween(StringX, "<" & sStartTag & ">", sEndTag)
    sTempString = sTempString & sDataString
    CountRootTags = CountRootTags + 1
    iPos = InStr(1, sEndTag, StringX, vbTextCompare) + Len(sEndTag)
    ' InStr(Text, Textstring1)
    'mvarChildString = GetInBetween(mvarWholeString, "<" & GetFullTag(mvarWholeString) & ">", "</" & mvarTag & ">")
Wend



End Function
Private Function GetFullTag(ByVal StringX As String, Optional lStartPos As Long) As String
'This function will get the root tag name of the string from a given position , including the ATTRIBUTES

Dim ii As Integer
If lStartPos < 2 Then lStartPos = 1
If lStartPos > Len(StringX) Then lStartPos = Len(StringX)

For ii = lStartPos To Len(StringX)
    If Mid(StringX, ii, 1) = "<" And (ii > lStartPos + 2) Then
        Exit Function
    Else
        If Mid(StringX, ii, 1) = ">" Or Mid(StringX, ii, 1) = "/" Then
            GetFullTag = Mid(StringX, 2, ii - 2)
            Exit Function
        End If
    End If
Next ii
End Function
Private Function GetInBetween(ByVal Text As String, ByVal Textstring1 As String, ByVal Textstring2 As String) As String


    If InStr(Text, Textstring1) = 0 Or InStr(Text, Textstring2) = 0 Then
        GetInBetween = ""
        Exit Function
    End If
    GetInBetween = Mid(Text, InStr(Text, Textstring1) + Len(Textstring1), InStr(Text, Textstring2) - (Len(Textstring1) + InStr(Text, Textstring1)))
End Function


'Private Function GetTagByNumber(ByVal StringX As String, iNumber As Integer) As String
'Dim LastOpen As Integer
'Dim LastClose As intger
'
'For ii = 1 To Len(StringX)
'    If Mid(StringX, ii, 1) = ">" Then
'        LastClose = ii
'        If LastOpen = iNumber Then ' Housten, we have a match
'            GetTagByNumber = Mid(StringX, LastOpen, (LastClose - LastOpen))
'        End If
'    End If
'    If Mid(StringX, ii, 1) = "<" Then
'        LastOpen = ii
'    End If
'
'Next ii
'
'
'End Function

Public Function BuildXML(ByVal sXML As String, oTreeView As Object) As Boolean
'Populate the Objects and Treeview
Dim NodX As Node


mvarWholeString = sXML ' The Whole String
mvarTag = GetTagName(mvarWholeString) 'This objects tag name
mvarChildString = GetInBetween(mvarWholeString, "<" & GetFullTag(mvarWholeString) & ">", "</" & mvarTag & ">") ' The Inner String of the element
Set mvarTreeView = oTreeView ' Set the treeview to build


'oTreeView.Nodes.Clear
'Set NodX = oTreeView.Nodes.Add(, , mvarTag, mvarTag, 1) ' Parent
NodeCount = NodeCount + 1
Set NodX = oTreeView.Nodes.Add(mvarRootTag, tvwChild, mvarTag, mvarTag, 1)

bCheck = BuildAttributes(GetFullTag(mvarChildString))
bCheck = BuildElements(mvarChildString, oTreeView)

End Function
Public Function BuildAttributes(ByVal StringX As String) As Boolean
'Build the Attributes

MsgBox StringX

End Function

Public Function BuildElements(ByVal StringX As String, oTreeView As Object) As Boolean

'Build the Elements contained in the InnerHTMLstring
Dim sTagString As String
Dim sInnerTag As String
Dim ii As Integer
'For ii = 1 To CharCount(StringX, "<")
'    GetTagByNumber
'Next ii

'sTagString = GetTagName(StringX)

For ii = 1 To CountRootTags(StringX)
'    Dim objXML2 As New clsXML
'    'objXML.BuildXML
'    objXML2.RootTag = mvarTag
'    mvarElements.Add objXML2
'
'    'sInnerTag = GetInBetween(mvarWholeString, "<" & GetFullTag(StringX) & ">", "</" & sTagString & ">") ' The Inner String of the element
'    bCheck = objXML2.BuildXML(mvarChildString, oTreeView)
'    Set objXML2 = Nothing
'    MsgBox StringX
Next ii
End Function
'Private Function CharCount(ByVal StringX As String, ByVal CharX As String) As Long
''Counts the occurence of a character in a string
'
'Dim ii As Long
'For ii = 1 To Len(CharX)
'    If Mid(StringX, ii, 1) = CharX Then
'        CharCount = CharCount + 1
'    End If
'Next ii
'
'End Function
