VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "objTreeBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim bCheck As Boolean
Dim NodX As Node
Dim NewObj As Object


Public Function BuildTree(objXML As Object, TreeX As TreeView, ParentName As String, MasterObj As Object) As Boolean
'This will start the building of the tree into however many levels are required
'On Error GoTo error_h
On Error Resume Next
Dim ii As Integer
Dim lAttKey As String
Dim lEleKey As String

 ' Build the Attributes
    For ii = 0 To objXML.Attributes.length - 1
            If objXML.Attributes.Item(ii).nodeName <> "UnID_trv_node_val_22134" Then
                lAttKey = GetKey
                Set NodX = TreeX.Nodes.Add(ParentName, tvwChild, lAttKey, objXML.Attributes.Item(ii).nodeName, 4) ' Child
                Set NodX = TreeX.Nodes.Add(lAttKey, tvwChild, GetKey, objXML.Attributes.Item(ii).Text, 5)  ' Child
            End If
    Next ii

' Build the Elements
    For ii = 0 To objXML.childNodes.length - 1
        lEleKey = GetKey
        If objXML.childNodes.Item(ii).nodeType = NODE_TEXT Then
            Set NodX = TreeX.Nodes.Add(ParentName, tvwChild, lEleKey, objXML.childNodes.Item(ii).Text, 6) ' Child
            Set NewObj = MasterObj.createAttribute("UnID_trv_node_val_22134")
            NewObj.Value = lEleKey
            objXML.childNodes(ii).Attributes.setNamedItem NewObj
        Else
            Set NodX = TreeX.Nodes.Add(ParentName, tvwChild, lEleKey, objXML.childNodes.Item(ii).nodeName, 3) ' Child
            Set NewObj = MasterObj.createAttribute("UnID_trv_node_val_22134")
            NewObj.Value = lEleKey
            objXML.childNodes(ii).Attributes.setNamedItem NewObj
            Dim objXMLchild As New objTreeBuild
            bCheck = objXMLchild.BuildTree(objXML.childNodes.Item(ii), TreeX, lEleKey, MasterObj)
            Set objXMLchild = Nothing
        End If
    Next ii
    
   
    
    Exit Function
error_h:
    MsgBox Err.Description
End Function

