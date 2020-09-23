VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmInfoTree 
   BackColor       =   &H00F0EFE3&
   Caption         =   "Info Tree"
   ClientHeight    =   5745
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4620
   Icon            =   "frmInfoTree.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   4620
   Begin MSComctlLib.TreeView TreeView 
      Height          =   2340
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4128
      _Version        =   393217
      Indentation     =   441
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   6
      Appearance      =   1
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Default         =   -1  'True
      Height          =   345
      Left            =   3630
      TabIndex        =   3
      Top             =   5325
      Width           =   810
   End
   Begin VB.TextBox txtPass 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2505
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5325
      Width           =   1050
   End
   Begin VB.Label lblPass 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   2
      Top             =   5325
      Width           =   2445
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&Info"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Changes"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSepBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel All Changes"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo to last Deleted Tree"
      End
      Begin VB.Menu mnuSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContract 
         Caption         =   "C&ontract All (Diminish)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuExpand 
         Caption         =   "&Expand All"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuShowKey 
         Caption         =   "Show Key Values"
      End
      Begin VB.Menu mnuSortTree 
         Caption         =   "Sort Tree &Alphabetically"
         Checked         =   -1  'True
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSepBar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddStats 
         Caption         =   "Add Statistics to InfoTree"
      End
      Begin VB.Menu mnuAddHelp 
         Caption         =   "Add Help to InfoTree"
      End
      Begin VB.Menu mnuSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit (Saves Changes)"
      End
   End
   Begin VB.Menu mnuBranch 
      Caption         =   "&Branch"
      Begin VB.Menu mnuMove 
         Caption         =   "&Move"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuBranchContract 
         Caption         =   "C&ontract"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuBranchExpand 
         Caption         =   "&Expand"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuSepBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuHTML 
         Caption         =   "Publish as HTML"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "*Export to Text File"
      End
      Begin VB.Menu mnuSepBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteBranch 
         Caption         =   "&Delete Branch"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuNode 
      Caption         =   "&Node"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuEditNode 
         Caption         =   "&Edit (or Dbl-Click)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "*&Replace"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuDumpTag 
         Caption         =   "&Dump Tag (Debug only)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuPassword 
         Caption         =   "Add/Remove &Password"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSepbar0303 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGotoLink 
         Caption         =   "&Goto Link"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuMakeLink 
         Caption         =   "&Link This Node"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuSepBar9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNodeBold 
         Caption         =   "&Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuNodeSorted 
         Caption         =   "&Sorted"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSepBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteNode 
         Caption         =   "Delete (Children remain)"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "frmInfoTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Form Needs:
' COMPONENT: Microsoft Windows Common Controls 6.0 (SP3)
'            (MSCOMCTL.OCX)
' REFERENCE: Microsoft DAO 3.5 Object Library (or later)
'            (DAO350.DLL)
' TREE2.ICO  - Icon for the Form.
'

Dim ITws As Workspace
Dim ITdb As Database
Dim ITrst As Recordset
Dim ITConfirmPassword As String
Dim ITFind As String
Dim ITNextNodeID As Long
Dim ITPasswordApproved As Date
Dim ITPrintLevel As Integer

'Window backcolor is &H00F0EFE3&
Const vbDarkRed = &H80&
'Const vbBlackRed = &H40&
'Const vbMediumOrange = &H40C0&
'Const vbDarkOrange = &H4080&
'Const vbMediumPurple = &HC000C0&
Const vbDarkGreen = &H8000&
Const vbDarkPurple = &H800080


Const KEYSYMBOL = "   |-->"
'These are predefined Node Names that can be
'added to the tree via Menu Options.
Const ABOUT_NODE = "About InfoTree"
Const STATS_NODE = "InfoTree Stats"
Const STATS_LAST_UPDATED = "Date Last Updated"
Const STATS_NODE_COUNT = "Node Count"
Const HELP_NODE = "InfoTree Help"

Private Function Open_Database() As Boolean
Dim AddInfoNode As Boolean
Dim DbFilename$
Dim DbFilePath$
Dim Ix As Index
Dim ITfd As Field
Dim rst As Recordset
Dim SQL$
Dim tdf As TableDef

Dim INFOTREE_DBFILEPATHandNAME As String
'This Path MUST Already exist (this form will create the .mdb file)
INFOTREE_DBFILEPATHandNAME = RTrim$(App.Path) & "\InfoTree.itr"

  Set ITws = DBEngine.Workspaces(0)   ' Use Main Program's Workspace
  DbFilePath$ = StripFile(INFOTREE_DBFILEPATHandNAME) ' DataBase Location
  DbFilename$ = INFOTREE_DBFILEPATHandNAME            ' DataBase Name
  If Len(Dir(DbFilePath, vbDirectory)) = 0 Then
    MsgBox DbFilePath, vbExclamation + vbOKOnly, "Sorry, Path does NOT exist!"
    Exit Function
  End If
  If Not ITFileExists(DbFilename$) Then
    MsgBox "Creating Database " & DbFilename$
    Set ITdb = ITws.CreateDatabase(DbFilename$, dbLangGeneral)
  Else
    Set ITdb = ITws.OpenDatabase(DbFilename$)
  End If

  If Not ITTableDefExists("InfoTree") Then
    Set tdf = ITdb.CreateTableDef("InfoTree")     ' Create TableDef
    MsgBox "Creating Table: " & tdf.Name
    Set ITfd = tdf.CreateField("IID", dbLong)     ' Define a Field
    ITfd.Attributes = dbAutoIncrField
    tdf.Fields.Append ITfd                        ' Add field to TableDef
    Set ITfd = tdf.CreateField("ParentNode", dbText, 10) ' Define a Field
    ITfd.AllowZeroLength = True
    tdf.Fields.Append ITfd                        ' Add field to TableDef
    Set ITfd = tdf.CreateField("NodeID", dbText, 10)   ' Define a Field
    tdf.Fields.Append ITfd                        ' Add field to TableDef
    Set ITfd = tdf.CreateField("NodeName", dbText, 255) ' Define a Field
    ITfd.AllowZeroLength = True
    tdf.Fields.Append ITfd                        ' Add field to TableDef
    Set ITfd = tdf.CreateField("Password", dbText, 20) ' Define a Field
    ITfd.AllowZeroLength = True
    tdf.Fields.Append ITfd                        ' Add field to TableDef
    Set ITfd = tdf.CreateField("Bold", dbBoolean) ' Define a Field
    tdf.Fields.Append ITfd                        ' Add field to TableDef
    Set ITfd = tdf.CreateField("OnTree", dbBoolean)  ' Define a Field
    tdf.Fields.Append ITfd                        ' Add field to TableDef
    Set ITfd = tdf.CreateField("LinksToNode", dbText, 10) ' Define a Field
    ITfd.AllowZeroLength = True
    tdf.Fields.Append ITfd                        ' Add field to TableDef
    Set ITfd = tdf.CreateField("NLV", dbBoolean)  ' Define a Field
    tdf.Fields.Append ITfd                        ' Add field to TableDef
    ITdb.TableDefs.Append tdf                     ' Append TableDef to Database
    AddInfoNode = True
  End If
  If Not ITKeyExists(ITdb.TableDefs("InfoTree").Indexes, "PrimaryKey") Then
    Set Ix = ITdb.TableDefs("InfoTree").CreateIndex("PrimaryKey")
    Ix.Primary = True       ' Field values are unique
    Set ITfd = Ix.CreateField("IID")
    Ix.Fields.Append ITfd
    ITdb.TableDefs("InfoTree").Indexes.Append Ix
  End If
'  If Not ITKeyExists(ITdb.TableDefs("InfoTree").Indexes, "NodeKey") Then
'    Set Ix = ITdb.TableDefs("InfoTree").CreateIndex("NodeKey")
'    Set ITfd = Ix.CreateField("ParentNode")
'    Ix.Fields.Append ITfd
'    Set ITfd = Ix.CreateField("NodeName")
'    Ix.Fields.Append ITfd
'    Ix.Primary = True        ' Nodes must be unique
'    ITdb.TableDefs("InfoTree").Indexes.Append Ix
'  End If
  
  SQL$ = "SELECT * from InfoTree " & _
         "WHERE InfoTree.NLV=False " & _
         "Order by InfoTree.ParentNode, InfoTree.NodeID;"
  Set ITrst = ITdb.OpenRecordset(SQL$, dbOpenDynaset)
  If AddInfoNode Then
    With ITrst
      .AddNew
        !ParentNode = ""
        !NodeID = "K00001"
        !NodeName = "Info"
      .Update
    End With
  End If
  Load_Treeview
  If AddInfoNode Then
    mnuAddStats_Click
    mnuAddHelp_Click
  End If
  Open_Database = True
End Function

Private Function Load_Treeview() As Boolean
Dim AllNodesOnTree As Boolean
Dim NodeText As String
Dim IsOnTree As Boolean
Dim ThisNode As Node
  TreeView.Nodes.Clear
  With ITrst
    If Not (.BOF And .EOF) Then
      .MoveFirst
      'First, mark all database records as Off the tree
      Do Until .EOF
        .Edit
          !OnTree = False
        .Update
        'While looping, calculate highest NodeID
        ITNextNodeID = ITMax(ITNextNodeID, _
          CLng(Right(!NodeID, Len(!NodeID) - 1)) + 1)
        .MoveNext
      Loop
      Do Until AllNodesOnTree
        'Because you can't add a child node to a treeview if
        'the parent node hasn't been added, you must
        'recursively pass the file until all nodes are added.
        AllNodesOnTree = True
        .MoveFirst
        Do Until .EOF
          If Not !OnTree Then
            NodeText = !NodeName
            If mnuShowKey.Checked Then
              NodeText = NodeText & KEYSYMBOL & !NodeID
            End If
            .Edit
              !OnTree = AddNode(!ParentNode, !NodeID, NodeText)
              If !OnTree Then
                Set ThisNode = TreeView.Nodes(CStr(!NodeID))
'If ThisNode.Key = "K00980" Then Stop
                ThisNode.Tag = "|IID" & !NodeID & "|"
                If Len(!Password) > 0 Then
                  ThisNode.Tag = ThisNode.Tag & "|PW" & !Password & "|"
                End If
                If Len(!LinksToNode) > 0 Then
                  ThisNode.Tag = ThisNode.Tag & "|LNK" & !LinksToNode & "|"
                End If
                If !Bold Then ThisNode.Bold = !Bold
                Set_Node_Colors ThisNode
              End If
              AllNodesOnTree = AllNodesOnTree And !OnTree
            .Update
            End If
          .MoveNext
        Loop
        DoEvents    'Just to give User a chance to close.
      Loop
      TreeView.Nodes(1).Expanded = True
    End If
  End With
  TreeIsDirty False       'By Definition
End Function

Private Function Save_Treeview() As Boolean
Dim ThisNode As Node
Dim NodeKey As String
Dim NodePrpKey As String
Dim NodePrpNameKey As String
Dim NodePrpValKey As String
Dim NodeText As String
Dim P As Integer
Dim SQL As String

  'Mark entire Table as No Longer Valid
  ' Step 1, erase all NLV records
  ' Step 2, mark all remaining records NLV
  SQL = "Delete * from InfoTree WHERE NLV=True;"
  ITdb.Execute SQL
  DoEvents: DoEvents
  With ITrst
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'      Do Until .EOF
'        If !NLV Then
'          .Delete
'          .MoveFirst
'        Else
'          .MoveNext
'        End If
'      Loop
'    End If
    .Requery
    If Not (.BOF And .EOF) Then
      .MoveFirst
      Do Until .EOF
        .Edit
          !NLV = True
        .Update
        .MoveNext
      Loop
    End If
  End With
  
  'Now pass all branches of Treeview, save to InfoTree Table
  With TreeView
    If .Nodes.Count > 0 Then
      For Each ThisNode In .Nodes
        If ThisNode.ForeColor = vbBlue Then
          ThisNode.ForeColor = vbBlack
        End If
        Do Until InStr(1, ThisNode.Tag, "|ND|") = 0
          ThisNode.Tag = StripFromTag(ThisNode.Tag, "ND")
        Loop
        ITrst.AddNew          'Add a record for this node
          NodeText = ThisNode.Text
          If Not (ThisNode.Parent Is Nothing) Then
            ITrst!ParentNode = ThisNode.Parent.Key
            If NodeText = STATS_LAST_UPDATED Then
              If ThisNode.Parent.Text = STATS_NODE Then
                If Not (ThisNode.Child Is Nothing) Then
                  ThisNode.Child.Text = Format(Date)
                End If
              End If
            End If
            If NodeText = STATS_NODE_COUNT Then
              If ThisNode.Parent.Text = STATS_NODE Then
                If Not (ThisNode.Child Is Nothing) Then
                  ThisNode.Child.Text = .Nodes.Count
                End If
              End If
            End If
          Else
            ITrst!ParentNode = ""   'Top Level
          End If
          If mnuShowKey.Checked = True Then
            'Strip Key from end of Text, if needed.
            P = InStr(1, NodeText, KEYSYMBOL)
            If P > 0 Then
              NodeText = Left$(NodeText, P - 1)
            End If
          End If
          ITrst!NodeName = NodeText
          ITrst!NodeID = ThisNode.Key
          ITrst!Bold = ThisNode.Bold
          P = InStr(1, ThisNode.Tag, "|LNK")
          If P > 0 Then
            ITrst!LinksToNode = Mid(ThisNode.Tag, P + 4, InStr(P + 1, ThisNode.Tag, "|") - (P + 4))
          End If
          P = InStr(1, ThisNode.Tag, "|PW")
          If P > 0 Then
            ITrst!Password = Mid(ThisNode.Tag, P + 3, InStr(P + 1, ThisNode.Tag, "|") - (P + 3))
          End If
        ITrst.Update
      Next
      ITrst.Requery
    End If
    
  End With
  TreeIsDirty False
  Save_Treeview = True
End Function

Private Function ContractChildrenOf(ByVal FromNodeKey As String)  'Recursive Expand
Dim FromNode As Node
Dim ChildNode As Node
Dim SiblingNode As Node
  Set FromNode = TreeView.Nodes(FromNodeKey)
  If (FromNode Is Nothing) Then Exit Function
  Set ChildNode = FromNode.Child
  If (ChildNode Is Nothing) Then Exit Function  'No children
  
  'Expand this Child
  ChildNode.Expanded = False
  ExpandChildrenOf ChildNode.Key
  Set SiblingNode = ChildNode.FirstSibling
  If FromNode.Children > 1 Then
    If SiblingNode.Key = ChildNode.Key Then
      Set SiblingNode = SiblingNode.Next
    End If
  End If
  If Not (SiblingNode Is Nothing) Then  'If child has siblings
    Do Until SiblingNode.Key = ChildNode.Key
      'Expand this Sibling
      SiblingNode.Expanded = False
      ExpandChildrenOf SiblingNode.Key
      Set SiblingNode = SiblingNode.Next
      If (SiblingNode Is Nothing) Then Exit Do
    Loop
  End If
End Function

Private Function ExpandChildrenOf(ByVal FromNodeKey As String)  'Recursive Expand
Dim FromNode As Node
Dim ChildNode As Node
Dim SiblingNode As Node
  Set FromNode = TreeView.Nodes(FromNodeKey)
  If (FromNode Is Nothing) Then Exit Function
  Set ChildNode = FromNode.Child
  If (ChildNode Is Nothing) Then Exit Function  'No children
  
  'Expand this Child
  If InStr(1, ChildNode.Tag, "|PW") = 0 Then
    ChildNode.Expanded = True
    ExpandChildrenOf ChildNode.Key
  End If
  Set SiblingNode = ChildNode.FirstSibling
  If FromNode.Children > 1 Then
    If SiblingNode.Key = ChildNode.Key Then
      Set SiblingNode = SiblingNode.Next
    End If
  End If
  If Not (SiblingNode Is Nothing) Then  'If child has siblings
    Do Until SiblingNode.Key = ChildNode.Key
      'Expand this Sibling
      If InStr(1, SiblingNode.Tag, "|PW") = 0 Then
        SiblingNode.Expanded = True
        ExpandChildrenOf SiblingNode.Key
      End If
      Set SiblingNode = SiblingNode.Next
      If (SiblingNode Is Nothing) Then Exit Do
    Loop
  End If
End Function

Private Function HTMLChildrenOf(ByVal FromNodeKey As String, _
    ByVal CHNL As Integer)  'Recursive HTML List
Dim FontsBeg As String
Dim FontsEnd As String
Dim FromNode As Node
Dim ChildNode As Node
Dim SiblingNode As Node
  Set FromNode = TreeView.Nodes(FromNodeKey)
  If (FromNode Is Nothing) Then Exit Function
  Set ChildNode = FromNode.Child
  If (ChildNode Is Nothing) Then Exit Function  'No children
  
  FontsBeg = "<LI>"
  FontsEnd = "</LI>"
  If ChildNode.Bold Then
    FontsBeg = FontsBeg & "<B>"
    FontsEnd = "</B>" & FontsEnd
  End If
  'Print this Child
  Print #CHNL, FontsBeg & ChildNode.Text & FontsEnd
  Print #CHNL, "<UL>"
  HTMLChildrenOf ChildNode.Key, CHNL
  Print #CHNL, "</UL>"
  
  Set SiblingNode = ChildNode.FirstSibling
  If FromNode.Children > 1 Then
    If SiblingNode.Key = ChildNode.Key Then
      Set SiblingNode = SiblingNode.Next
    End If
  End If
  If Not (SiblingNode Is Nothing) Then  'If child has siblings
    Do Until SiblingNode.Key = ChildNode.Key
      'Print this Sibling
      FontsBeg = "<LI>"
      FontsEnd = "</LI>"
      If SiblingNode.Bold Then
        FontsBeg = FontsBeg & "<B>"
        FontsEnd = "</B>" & FontsEnd
      End If
      Print #CHNL, FontsBeg & SiblingNode.Text & FontsEnd
      Print #CHNL, "<UL>"
      HTMLChildrenOf SiblingNode.Key, CHNL
      Print #CHNL, "</UL>"
      Set SiblingNode = SiblingNode.Next
      If (SiblingNode Is Nothing) Then Exit Do
    Loop
  End If
End Function




'*********'*********'*********'*********!*********'*********'*********'**(79)**
'Parameters for Get_Environ (not case sensitive):
'TMP (DOS Temp Directory), TEMP (Windows Temp Directory), PROMPT (DOS Prompt),
'WINBOOTDIR (usually C:\Windows), COMSPEC (Path&FileName of Command.com),
'WINDIR (Windows Directory), MAC (MAC Address), NAME (Network Login ID),
'CX (Novell Container), WINDOWS_LOGIN,  PATH
Private Function ITGet_Environ(ByVal iStr As String) As String
Dim str As String
Dim X As Integer
  'Note:  Environ can also be called like this: Environ("NAME")
  str = UCase(Trim$(iStr))
  For X = 1 To 30
    If InStr(1, UCase(Environ(X)), str & "=") Then
      'When accessed by Number, returns PATH=....
      ITGet_Environ = Mid$(Environ(X), (Len(str) + 2))
      Exit For
    End If
  Next X
End Function
Private Function ITFileExists(ByVal File$) As Boolean
On Error Resume Next
  If FileLen(File$) > -1 Then
    If Err = 0 Then
      ITFileExists = True
    End If
  End If
End Function
'Determines whether the program's running from VB's IDE or an Executable
'Let's you put extra data on the screen that's only available while testing, not in the finished product. (Note. This is also a property of Rich's Bug.Frm Error-handler. If you're going to use that, you needn't duplicate the code here, just use Bug.InIDE).
Public Function ITIn_IDE() As Boolean
    On Error GoTo ErrHand
    ' We are assuming the program is run from an .exe
    ITIn_IDE = False
    ' An error will occur if this is not bypassed.
    Debug.Print "Two" + 2
    ' Otherwise, we know this is running from the .exe
    Exit Function
ErrHand:
    ITIn_IDE = True
End Function

Private Function ITKeyExists(Client As Object, Key As String) As Boolean
' ITKeyExists
' Returns whether an item is part of a collection
'
' Client - pointer to collection
' Key    - unique key of object being tested
' Returns: True if item with Key exists, False if item doesn't exist
    On Error GoTo ITKeyExistsErrorHandler
    Dim Obj As Object
    Set Obj = Client.Item(Key)
    Set Obj = Nothing
    ITKeyExists = True
Exit Function
ITKeyExistsErrorHandler:
    ITKeyExists = False
End Function
Private Function ITMax(A As Variant, B As Variant)
     If (A > B) Then ITMax = A Else ITMax = B
End Function
Private Function ITSetMid(ByVal istrToModify As String, ByVal istrToInsert As String, _
  ByVal ilngStart As Long, ByVal ilngCharsToRemove As Long) As String
'Inputs:  istrToModify - The string you wish to modify.
'         istrToInsert - The string you are inserting into istrToModify.
'         ilngStart - The character In istrToModify from which istrToInsert must be inserted.
'         ilngCharsToRemove - How many characters istrToInsert must replace in
'           istrToModify, from ilngStart.
'           - Set this To 0,  istrToInsert will be inserted without over-writing any characters in istrToModify.
'           - Set this To -1, istrToInsert will over-write to the length of istrToInsert
'Returns: An appropriate merge of istrToModify and istrToInsert.
'
    Dim first As String
    Dim last As String
    If ilngCharsToRemove < 0 Then ilngCharsToRemove = Len(istrToInsert)
    first = Left(istrToModify, ilngStart - 1)
    last = Right(istrToModify, Len(istrToModify) - (ilngStart + ilngCharsToRemove) + 1)
    ITSetMid = first & istrToInsert & last
End Function
Private Function ITTableDefExists(ByVal iTableName As String) As Boolean
'Does Table Exist in Database?
Dim X As Integer
  ITTableDefExists = False
  For X = 0 To ITdb.TableDefs.Count - 1
    If ITdb.TableDefs(X).Name = iTableName Then
      ITTableDefExists = True
      Exit For
    End If
  Next
End Function

'Private Sub cmdSave_Click()
'Dim Selected_Node As Node
'Dim NodeKey As String
'Dim ParentNode As String
''Dim str As String
'  If Len(txtNodeName.Text) = 0 Then
'    If Len(txtPrpName.Text) = 0 Then
'      cmdCancel_Click
'      Exit Sub
'    End If
'  End If
'  Set Selected_Node = TreeView.SelectedItem
'  If (Selected_Node Is Nothing) Then
'    ParentNode = ""
'    NodeKey = ""                    'Create a new Node
'  Else
'    ParentNode = Selected_Node.Key
'    If Len(txtNodeName.Text) = 0 Then
'      'You left the node blank, add to selected existing node.
'      txtNodeName.Text = Selected_Node.Text
'      DoEvents
'      ParentNode = Selected_Node.Parent.Key
'      NodeKey = Selected_Node.Key   'Use Existing node
'    End If
'  End If
'
'  txtNodeName.Text = Trim$(Left$(txtNodeName.Text, 255))
'  txtPrpName.Text = Trim$(Left$(txtPrpName.Text, 255))
'  txtPrpVal.Text = Trim$(Left$(txtPrpVal.Text, 255))
'
'  Select Case ITMode
'  Case 1            'Add New Node
'    AddNode ParentNode, NodeKey, txtNodeName.Text, txtPrpName.Text, txtPrpVal.Text
'  Case 2            'Edit Node
'
'  End Select
'  txtNodeName.Text = ""
'  txtPrpName.Text = ""
'  txtPrpVal.Text = ""
'  Cover_Edit_Boxes
'  TreeIsDirty True  'Save changes to Database on exit
'  ITMode = 0        'Back to View Mode
'End Sub
'
'Private Function AddNode(ByVal ParentNodeKey As String, _
'    NodeID As String, _
'    NodeName As String, _
'    PrpName As Variant, _
'    PrpVal As Variant) As Boolean
'Dim NodeKey As String
'Dim NodePrpNameKey As String
'Dim NodePrpValKey As String
'  If Len(NodeID) = 0 Then
'    NodeID = "K" & Format(ITNextNodeID, "00000")
'    ITNextNodeID = ITNextNodeID + 1
'  End If
'  PrpName = "" & PrpName
'  PrpVal = "" & PrpVal
'  NodeKey = NodeID
'  With TreeView
'    If Not ITKeyExists(TreeView.Nodes, NodeKey) Then
'      'Add Node if needed
'      If Len(ParentNodeKey) = 0 Then
'        TreeView.Nodes.Add , tvwFirst, NodeKey, NodeName  'Top Level
'      Else
'        On Error GoTo Trap_No_Parent
'        TreeView.Nodes.Add ParentNodeKey, tvwChild, NodeKey, NodeName
'        On Error GoTo 0
'      End If
'    End If
'    If Len(PrpName) > 0 Then
'      NodePrpNameKey = "K" & Format(ITNextNodeID, "00000")
'      ITNextNodeID = ITNextNodeID + 1
'      NodePrpValKey = "K" & Format(ITNextNodeID, "00000")
'      ITNextNodeID = ITNextNodeID + 1
'      'Add Property
'      TreeView.Nodes.Add NodeKey, tvwChild, NodePrpNameKey, PrpName
'      TreeView.Nodes.Add NodePrpNameKey, tvwChild, NodePrpValKey, PrpVal
'    End If
'  End With
'  AddNode = True
'  TreeIsDirty True
'Exit_Function:
'Exit Function
'Trap_No_Parent:
'  If Err.Number = 35601 Then
'    'Node was added to a non-existent parent.
'    'This happens frequently in Load_Treeview, while it
'    '   recursively attaches nodes to the tree.
'    'Let the Function return False.
'    Resume Exit_Function
'  Else
'    On Error GoTo 0
'  End If
'Resume
'End Function

Private Function KeyForText(ByVal iText As String) As String
Dim ThisNode As Node
  With TreeView
    For Each ThisNode In .Nodes
      If ThisNode.Text = iText Then
        KeyForText = ThisNode.Key
        Exit For
      End If
    Next
  End With
End Function

'Same as AddNode, but turns new node blue
Private Function NewNode(ByVal ParentNodeKey As String, _
    ByRef NodeKey As String, _
    NodeName As String) As Boolean
Dim NewKey As String
  If Len(NodeKey) > 0 Then NewKey = NodeKey
  NewNode = AddNode(ParentNodeKey, NewKey, NodeName)
  TreeIsDirty True
  If NewNode Then
    TreeView.Nodes(NewKey).ForeColor = vbBlue
    NodeKey = NewKey
  End If
End Function

Private Function AddNode(ByVal ParentNodeKey As String, _
    ByRef NodeKey As String, _
    NodeName As String) As Boolean
  If Len(NodeKey) = 0 Then
    NodeKey = NextFreeNodeID
  End If
  With TreeView
    If Not ITKeyExists(TreeView.Nodes, NodeKey) Then
      'Add Node if needed
      If Len(ParentNodeKey) = 0 Then
        Select Case NodeName
        Case "Info"
          .Nodes.Add , tvwFirst, NodeKey, NodeName 'Top Level
        Case ABOUT_NODE
          DoEvents
          .Nodes.Add .Nodes(1).Key, tvwLast, NodeKey, NodeName  'Top Level
        Case Else
          DoEvents
          .Nodes.Add .Nodes(1).Key, tvwNext, NodeKey, NodeName  'Top Level
        End Select
      Else
        On Error GoTo Trap_No_Parent
        .Nodes.Add ParentNodeKey, tvwChild, NodeKey, NodeName
        On Error GoTo 0
      End If
      If mnuSortTree.Checked Then
        TreeView.Nodes(NodeKey).Sorted = True
      End If
    End If
  End With
  AddNode = True
  'CAN'T SAY THIS HERE! TreeIsDirty True
Exit_Function:
Exit Function
Trap_No_Parent:
  If Err.Number = 35601 Then
    'Node was added to a non-existent parent.
    'This happens frequently in Load_Treeview, while it
    '   recursively attaches nodes to the tree.
    'Let the Function return False.
    Resume Exit_Function
  Else
    On Error GoTo 0
  End If
Resume
End Function

Private Sub cmdAccept_Click()
  If Len(txtPass.Text) > 0 Then   'Leave blank to stop function
    If mnuPassword.Caption = "Add &Password" Then
      If Len(ITConfirmPassword) = 0 Then
        ITConfirmPassword = txtPass.Text
        lblPass.Caption = "CONFIRM Password:"
        txtPass.Text = ""
        txtPass.SetFocus
        Exit Sub
      Else
        If txtPass.Text <> ITConfirmPassword Then
          MsgBox "Sorry, the 2 passwords don't match. Try Again."
          GoTo Restore_Treeview
        Else
          TreeView.Enabled = True
          With TreeView
            .SelectedItem.Tag = AddToTag(.SelectedItem.Tag, "PW", ITConfirmPassword)
            TreeIsDirty True
          End With
        End If
      End If
    Else  'Else Remove Password
      TreeView.Enabled = True
      With TreeView
        If InStr(1, UCase(.SelectedItem.Tag), "|PW" & UCase(txtPass.Text)) > 0 Then
          .SelectedItem.Tag = StripFromTag(.SelectedItem.Tag, "PW")
          TreeIsDirty True
        Else
          MsgBox "Sorry, that's not the correct Password. Try Again."
        End If
      End With
    End If
  End If
  
Restore_Treeview:
  TreeView.Enabled = True
  Form_Resize   'Pull Treeview back to bottom of form.
  Set_Node_Colors TreeView.SelectedItem
End Sub

Private Sub Form_Load()
Dim lngTop As Long
Dim lngLeft As Long
Dim lngHeight As Long
  lngTop = 100
  lngLeft = (Screen.Width - Me.Width) - 100
  lngHeight = Screen.Height - 1000
  Me.Move lngLeft, lngTop, Me.Width, lngHeight
  If Not ITIn_IDE Then
    'Debugging tool, don't show in final product
    mnuDumpTag.Visible = False
  End If
  Open_Database
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    TreeView.Width = Me.ScaleWidth - 120
    TreeView.Height = Me.ScaleHeight - 120
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If TreeIsDirty Then
    mnuSave_Click
  End If
  If Not (ITrst Is Nothing) Then
    ITrst.Close
  End If
  Set ITdb = Nothing
  Set ITws = Nothing
End Sub

Private Sub mnuAdd_Click()
Dim Added_Node As Node
Dim Parent_Node As Node
Dim AddedNodeID As String
Dim ParentNodeID As String
  Set Parent_Node = TreeView.SelectedItem
  If Not (Parent_Node Is Nothing) Then
    ParentNodeID = Parent_Node.Key
    Parent_Node.Sorted = False  'Add new node at bottom of list.
  End If
  AddedNodeID = NextFreeNodeID
  If NewNode(ParentNodeID, AddedNodeID, "") Then
    Set Added_Node = TreeView.Nodes(AddedNodeID)
    Added_Node.EnsureVisible
    Set TreeView.SelectedItem = Added_Node
    TreeView.StartLabelEdit
  End If
End Sub

Private Sub mnuAddHelp_Click()
Dim AboutKey As String
Dim HelpKey As String
Dim HelpTopicKey As String
Dim HelpSubTopicKey As String
Dim PropKey As String
  AboutKey = KeyForText(ABOUT_NODE)
  If Len(AboutKey) = 0 Then
    AddNode "", AboutKey, ABOUT_NODE
  End If
  NewNode AboutKey, HelpKey, HELP_NODE
    NewNode HelpKey, HelpTopicKey, "Menu Items"
      NewNode HelpTopicKey, HelpSubTopicKey, "Info"
        NewNode HelpSubTopicKey, PropKey, "Cancel All Changes"
          NewNode PropKey, "", "a) Re-reads the tree from the disk."
          NewNode PropKey, "", "b) All changes (blue text) in the tree will be lost."
          NewNode PropKey, "", "c) Note that the top-most node may be blue, but won't be deleted."
        PropKey = ""
        NewNode HelpSubTopicKey, PropKey, "Undo to Last Deleted Tree"
          NewNode PropKey, "", "a) Each time a save is performed, the tree on the disk is"
          NewNode PropKey, "", "b)   marked deleted, and the tree in memory is written to the disk."
          NewNode PropKey, "", "c) This option allows you to restore the tree on the disk that"
          NewNode PropKey, "", "d)   was marked deleted (all changes since then will be lost)."
        PropKey = ""
        NewNode HelpSubTopicKey, PropKey, "Exit"
          NewNode PropKey, "", "a) Will Save changes to Tree, then Exit."
          NewNode PropKey, "", "b) (Use 'Info>Cancel All Changes' menu option to abandon changes)."
          NewNode PropKey, "", "c) Using the X button for the window has the same effect."
        PropKey = ""
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Branch"
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Node"
        NewNode HelpSubTopicKey, PropKey, "Delete (Children Remain)"
          NewNode PropKey, "", "a) Removes a Node, But does NOT remove the children."
          NewNode PropKey, "", "b) (Children are re-attached to the parent of the deleted node)."
        PropKey = ""
      HelpSubTopicKey = ""
    HelpTopicKey = ""
  
    NewNode HelpKey, HelpTopicKey, "Siblings"
      NewNode HelpTopicKey, HelpSubTopicKey, "Defined"
        NewNode HelpSubTopicKey, "", "a) All nodes on the same line are Siblings."
        NewNode HelpSubTopicKey, "", "b) You could also say all Nodes that expand or Contract at the same time are Siblings."
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Copying All Siblings"
        NewNode HelpSubTopicKey, "", "a) Copy the Parent Node of the Siblings."
        NewNode HelpSubTopicKey, "", "b) Then use the 'Delete' Key to remove the parent node's name from the new location."
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Deleting All Siblings"
        NewNode HelpSubTopicKey, "", "a) Highlight the Parent Node, hit F2 then Ctrl-C."
        NewNode HelpSubTopicKey, "", "a) Delete the Parent Node of the Siblings."
        NewNode HelpSubTopicKey, "", "b) Hit F1 to create a Node, then Ctrl-V to paste the Parent node back in."
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Moving All Siblings"
        NewNode HelpSubTopicKey, "", "a) Move the Parent Node of the Siblings."
        NewNode HelpSubTopicKey, "", "b) Then use the 'Delete' Key to remove the parent node's name from the new location."
      HelpSubTopicKey = ""
    HelpTopicKey = ""
  
    NewNode HelpKey, HelpTopicKey, "Color Key"
      NewNode HelpTopicKey, HelpSubTopicKey, "Black on White"
        NewNode HelpSubTopicKey, "", "Standard Information, backed up on the disk."
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Blue on White"
        NewNode HelpSubTopicKey, "", "Data that has NOT been backed up to disk."
        NewNode HelpSubTopicKey, "", "Hit the Ctrl-S key to save this data."
        NewNode HelpSubTopicKey, "", "(It's automatically saved on Exit)."
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Green on Black"
        NewNode HelpSubTopicKey, "", "CORRUPTION! You should never see this."
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Green on White"
        NewNode HelpSubTopicKey, "", "Link to Another Node - Hit F11 to Jump there"
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Green on Yellow"
        NewNode HelpSubTopicKey, "", "A Node that's been flagged to have a link created."
        NewNode HelpSubTopicKey, "", "To remove, Click on the same node, then hit sF11."
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Magenta on White"
        NewNode HelpSubTopicKey, "", "Node has been Password-Protected."
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Red on Yellow"
        NewNode HelpSubTopicKey, "", "A Node that's been flagged to MOVE."
        NewNode HelpSubTopicKey, "", "To remove, Click on the same node, then hit F6."
      HelpSubTopicKey = ""
      NewNode HelpTopicKey, HelpSubTopicKey, "Blue on Yellow"
        NewNode HelpSubTopicKey, "", "A Node that's been flagged to COPY."
        NewNode HelpSubTopicKey, "", "To remove, Click on the same node, then hit F7."
      HelpSubTopicKey = ""
    HelpTopicKey = ""
  HelpKey = ""
End Sub

Private Sub mnuAddStats_Click()
Dim ValueKey As String
Dim PropKey As String
Dim StatsKey As String
Dim AboutKey As String
  If Len(KeyForText(ABOUT_NODE)) = 0 Then
    NewNode "", AboutKey, ABOUT_NODE
  Else
    AboutKey = KeyForText(ABOUT_NODE)
  End If
  NewNode AboutKey, StatsKey, STATS_NODE
  NewNode StatsKey, PropKey, "Date Tree Created"
  NewNode PropKey, ValueKey, Format(Date)
  PropKey = "": ValueKey = ""
  NewNode StatsKey, PropKey, STATS_LAST_UPDATED
  NewNode PropKey, ValueKey, Format(Date)
  PropKey = "": ValueKey = ""
  NewNode StatsKey, PropKey, STATS_NODE_COUNT
  NewNode PropKey, ValueKey, "<Unknown till Saved>"
End Sub

Private Sub mnuBranchContract_Click()
Dim TopNode As Node
Dim SiblingNode As Node
  Set TopNode = TreeView.SelectedItem
  If Not (TopNode Is Nothing) Then
    TopNode.Expanded = False
    ContractChildrenOf TopNode.Key
  End If
End Sub

Private Sub mnuBranchExpand_Click()
Dim TopNode As Node
Dim SiblingNode As Node
  Set TopNode = TreeView.SelectedItem
  If Not (TopNode Is Nothing) Then
    TopNode.Expanded = True
    ExpandChildrenOf TopNode.Key
  End If
End Sub

Private Sub mnuCancel_Click()
  Load_Treeview
  TreeIsDirty False
End Sub

Private Sub mnuContract_Click()
Dim i As Integer
  For i = 1 To TreeView.Nodes.Count ' goes threw each node and Collapses it
    TreeView.Nodes.Item(i).Expanded = False
  Next i
End Sub

Private Sub mnuCopy_Click()
Static From_Node As Node
Dim lngResult As Long
Dim To_Node As Node
Dim NewNodeID As String
  If mnuCopy.Caption = "&Copy" Then
    Set From_Node = TreeView.SelectedItem
    If (From_Node Is Nothing) Then
      MsgBox "Select the Node you wish to Copy, THEN hit F7", _
        vbOKOnly + vbInformation, "Can't Copy, no Node Selected"
      Exit Sub
    End If
'    If From_Node.Key = "K00001" Then
'      MsgBox "You can't copy the Highest Node," & vbCrLf & _
'        "there'd be nowhere to copy it TO.", _
'        vbOKOnly + vbInformation, "Can't Copy Top-Level node"
'      Exit Sub
'    End If
    mnuCopy.Caption = "Copying Node <" & Left$(From_Node.Text, 20) & ">"
    From_Node.Tag = AddToTag(From_Node.Tag, "CP", "")
    Set_Node_Colors From_Node
    Set TreeView.SelectedItem = From_Node.Parent
    Beep
  Else
    Set To_Node = TreeView.SelectedItem
    If (To_Node Is Nothing) Then
      lngResult = MsgBox("Select the Node you wish to Copy TO, THEN hit F7" & _
        vbCrLf & "  (or hit Cancel to stop the Copy)", _
        vbOKCancel + vbInformation, _
        "Copying Node <" & Left$(From_Node.Text, 20) & ">")
      If lngResult = vbCancel Then
        GoTo Cleanup_Copy
      Else
        Exit Sub
      End If
    End If
'    On Error GoTo Cleanup_Copy  'Handles Copying a node to itself
    'Copy the Selected Node to the new branch
    
    NewNodeID = NextFreeNodeID
    AddNode To_Node.Key, NewNodeID, From_Node.Text
    CopyChildrenOf From_Node.Key, NewNodeID, To_Node.Text 'Recursive copy
    TreeIsDirty True
Cleanup_Copy:
    From_Node.Tag = StripFromTag(From_Node.Tag, "CP")
    Set_Node_Colors From_Node
    Set From_Node = Nothing
    mnuCopy.Caption = "&Copy"
  End If
End Sub

Private Sub mnuDeleteBranch_Click()
Dim lngResult As Long
Dim Selected_Node As Node
Dim ChildNode As Node
  Set Selected_Node = TreeView.SelectedItem
  If (Selected_Node Is Nothing) Then
    MsgBox "You must highlight a node before deleting", _
      vbOKOnly + vbInformation, "Can't Delete Branch"
    Exit Sub
  End If
  If Selected_Node.Key = "K00001" Then
    lngResult = MsgBox("Are You sure you want to Delete the WHOLE tree?", _
      vbQuestion + vbYesNoCancel, "Confirm Deletion")
    If lngResult <> vbYes Then Exit Sub
  End If
  Set ChildNode = Selected_Node.Child
  If Not (ChildNode Is Nothing) Then
    lngResult = MsgBox("BRANCH: " & Selected_Node.Text, _
      vbQuestion + vbYesNoCancel, _
      "Confirm Branch Deletion")
    If lngResult <> vbYes Then Exit Sub
  End If
  TreeView.Nodes.Remove Selected_Node.Key
  TreeIsDirty True
End Sub

Private Sub mnuDeleteNode_Click()
Dim lngResult As Long
Dim Selected_Node As Node
Dim ChildNode As Node
Dim ParentKey As String
Dim ParentNode As Node
Dim SiblingNode As Node
  Set Selected_Node = TreeView.SelectedItem
  If (Selected_Node Is Nothing) Then
    MsgBox "You must highlight a node before deleting", _
      vbOKOnly + vbInformation, "Can't Delete Node"
    Exit Sub
  End If
  If Selected_Node.Key = "K00001" Then
    lngResult = MsgBox("Are You sure you want to Delete the Main Node?", vbQuestion + vbYesNoCancel, "Confirm Deletion")
    If lngResult <> vbYes Then Exit Sub
  End If
  Set ChildNode = Selected_Node.Child
  If Not (ChildNode Is Nothing) Then
    lngResult = MsgBox("NODE: " & Selected_Node.Text, _
      vbQuestion + vbYesNoCancel, _
      "Confirm Deletion (Children Remain)")
    If lngResult <> vbYes Then Exit Sub
    'Give all Node's children to Node's Parent
    Set ParentNode = Selected_Node.Parent
    If (ParentNode Is Nothing) Then
      ParentKey = ""
    Else
      ParentKey = ParentNode.Key
    End If
    Set SiblingNode = ChildNode.Next
    If SiblingNode.Key <> ChildNode.Key Then
      Do Until SiblingNode.Key = ChildNode.Key
        Set SiblingNode.Parent = ParentNode
        Set SiblingNode = ChildNode.Next
        If (SiblingNode Is Nothing) Then Exit Do
      Loop
    End If
    Set ChildNode.Parent = ParentNode
  End If
  TreeView.Nodes.Remove Selected_Node.Key
  TreeIsDirty True
End Sub

Private Sub mnuDumpTag_Click()
  MsgBox "<" & TreeView.SelectedItem.Tag & ">", _
    vbInformation + vbOKOnly, _
    "Tag for <" & Left$(TreeView.SelectedItem.Text, 14)
End Sub

Private Sub mnuEditNode_Click()
Static Edit_Node As Node
  Set Edit_Node = TreeView.SelectedItem
  If (Edit_Node Is Nothing) Then
    MsgBox "Select the Node you wish to Edit, THEN hit F2", _
      vbOKOnly + vbInformation, "Can't Edit, no Node Selected"
    Exit Sub
  Else
    With Edit_Node
      .EnsureVisible
      If Not (.Parent Is Nothing) Then
        .Parent.Sorted = False  'Changing text can 'Un-Sort'.
      End If
      TreeView.StartLabelEdit
    End With
  End If
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuExpand_Click()
Dim i As Integer
  For i = 1 To TreeView.Nodes.Count ' goes threw each node and Expandes it
    If InStr(1, TreeView.Nodes(i).Tag, "|PW") = 0 Then
      TreeView.Nodes.Item(i).Expanded = True
    End If
  Next i
End Sub

Private Function TreeIsDirty(Optional ByVal TF As Variant) As Boolean
Dim IsDirty As Boolean
Dim InfoNode As Node
  If Not IsMissing(TF) Then
    IsDirty = TF
    'mnuCancel is a global 'Is Dirty' flag,
    'but use TreeIsDirty to return the value.
    'Example:  If TreeIsDirty then Save_Treeview.
    mnuCancel.Enabled = IsDirty
    Set InfoNode = TreeView.Nodes("K00001")
    If Not (InfoNode Is Nothing) Then
      With InfoNode
        If IsDirty Then
          InfoNode.Tag = AddToTag(InfoNode.Tag, "ND") 'ND = Node Dirty
        Else
          InfoNode.Tag = StripFromTag(InfoNode.Tag, "ND")
        End If
        Set_Node_Colors InfoNode
      End With
    End If
  End If
  TreeIsDirty = mnuCancel.Enabled
  Set_Menu_Options
End Function

Private Sub mnuFind_Click()
Dim ChildNode As Node
Dim FoundFirstNode As Boolean
Dim ThisNode As Node

  ITFind = UCase(InputBox("Enter Text to search for:", "Find In Tree"))
  If Len(ITFind) = 0 Then Exit Sub

  mnuContract_Click
  'clear Found Colors from all branches
  For Each ThisNode In TreeView.Nodes
    If ThisNode.BackColor = vbCyan Then
      Set_Node_Colors ThisNode
    End If
  Next

  For Each ThisNode In TreeView.Nodes
    If InStr(1, UCase(ThisNode.Text), ITFind) > 0 Then
      ThisNode.EnsureVisible
      ThisNode.ForeColor = vbDarkPurple
      ThisNode.BackColor = vbCyan
      Set ChildNode = ThisNode.Child
      If Not (ChildNode Is Nothing) Then
        ThisNode.Child.EnsureVisible    'Expand the found node's Children
      End If
      If Not FoundFirstNode Then
        Set TreeView.SelectedItem = ThisNode
        FoundFirstNode = True
      End If
    End If
  Next
End Sub

Private Sub mnuGotoLink_Click()
Dim NodeTag As String
Dim NodeKey As String
Dim P As Integer
Dim ThisNode As Node
  Set ThisNode = TreeView.SelectedItem
  If Not (ThisNode Is Nothing) Then
    NodeTag = ThisNode.Tag
    P = InStr(1, NodeTag, "|LNK")
    If P > 0 Then
      NodeKey = Mid$(NodeTag, P + 4, 6)
      TreeView.SelectedItem = TreeView.Nodes(NodeKey)
      TreeView.SelectedItem.Selected = True
      TreeView.SelectedItem.EnsureVisible
    Else
      Beep
    End If
  Else
    Beep
  End If
End Sub

Private Sub mnuHTML_Click()
Dim CHNL As Integer
Dim File As String
Dim FontsBeg As String
Dim FontsEnd As String
Dim TopNode As Node
Dim SiblingNode As Node
  'If TreeIsDirty Then mnuSave_Click
  File = RegreadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Desktop")
  File = File & "\InfoTree.htm"
  CHNL = FreeFile
  Open File For Output As #CHNL
  Print #CHNL, "<HTML>"
  Print #CHNL, "<!Published by InfoTree on " & Format(Now()) & ">"
  Set TopNode = TreeView.SelectedItem
  If Not (TopNode Is Nothing) Then
    FontsBeg = ""
    FontsEnd = ""
    If TopNode.Bold Then
      FontsBeg = FontsBeg & "<B>"
      FontsEnd = "</B>" & FontsEnd
    End If
    Print #CHNL, FontsBeg & TopNode.Text & FontsEnd
    Print #CHNL, "<UL>"
    HTMLChildrenOf TopNode.Key, CHNL
    Print #CHNL, "</UL>"
  End If
  Print #CHNL, "</HTML>"
  Close #CHNL
  If ITFileExists(File) Then
    Shell ITGet_Environ("COMSPEC") & " /C " & Chr(34) & File & Chr(34)
  End If
  MsgBox "File <InfoTree.htm> is on your Desktop"
End Sub

Private Sub mnuMakeLink_Click()
Static From_Node As Node
Dim lngResult As Long
Dim To_Node As Node
  If mnuMakeLink.Caption = "&Link This Node" Then
    Set From_Node = TreeView.SelectedItem
    If (From_Node Is Nothing) Then
      MsgBox "Select the Node you wish to add a link to, THEN hit shift-F11", _
        vbOKOnly + vbInformation, "Can't Link, no Node Selected"
      Exit Sub
    End If
    If InStr(1, From_Node.Tag, "|LNK") > 0 Then
      lngResult = MsgBox("Node is already linked. OK to break old link?" & _
        vbCrLf & "  (or hit Cancel to break the Link without a new link)", _
        vbYesNoCancel + vbQuestion, _
        "Linking Node <" & Left$(From_Node.Text, 20) & ">")
      If lngResult = vbNo Then
        Set From_Node = Nothing
        Exit Sub
      ElseIf lngResult = vbCancel Then
        From_Node.Tag = StripFromTag(From_Node.Tag, "LNK")
        From_Node.Tag = AddToTag(From_Node.Tag, "ND") 'Node is Dirty
        TreeIsDirty True
        Set_Node_Colors From_Node
        Set From_Node = Nothing
        Exit Sub
      End If
      From_Node.Tag = StripFromTag(From_Node.Tag, "LNK")
      TreeIsDirty True
    End If
    mnuMakeLink.Caption = "Linking Node <" & Left$(From_Node.Text, 20) & ">"
    From_Node.Tag = AddToTag(From_Node.Tag, "NL")
    Set_Node_Colors From_Node
    Set TreeView.SelectedItem = From_Node.Parent
    Beep
  Else
    Set To_Node = TreeView.SelectedItem
    If (To_Node Is Nothing) Then
      lngResult = MsgBox("Select the Node you wish to Link TO, THEN hit s-F11" & _
        vbCrLf & "  (or hit Cancel to stop the Link)", _
        vbOKCancel + vbInformation, _
        "Linking Node <" & Left$(From_Node.Text, 20) & ">")
    Else
      From_Node.Tag = StripFromTag(From_Node.Tag, "NL")
      From_Node.Tag = AddToTag(From_Node.Tag, "LNK", To_Node.Key)
      TreeIsDirty True
    End If
    Set_Node_Colors From_Node
    Set From_Node = Nothing
    mnuMakeLink.Caption = "&Link This Node"
  End If
End Sub

Private Sub mnuMove_Click()
Static From_Node As Node
Dim lngResult As Long
Dim To_Node As Node
  If mnuMove.Caption = "&Move" Then
    Set From_Node = TreeView.SelectedItem
    If (From_Node Is Nothing) Then
      MsgBox "Select the Node you wish to Move, THEN hit F6", _
        vbOKOnly + vbInformation, "Can't Move, no Node Selected"
      Exit Sub
    End If
    If From_Node.Key = "K00001" Then
      MsgBox "You can't move the Highest Node," & vbCrLf & _
        "there'd be nowhere to move it TO.", _
        vbOKOnly + vbInformation, "Can't Move Top-Level node"
      Exit Sub
    End If
    mnuMove.Caption = "Moving Node <" & Left$(From_Node.Text, 20) & ">"
    From_Node.Tag = AddToTag(From_Node.Tag, "MV")
    Set_Node_Colors From_Node
    Set TreeView.SelectedItem = From_Node.Parent
    Beep
  Else
    Set To_Node = TreeView.SelectedItem
    If (To_Node Is Nothing) Then
      lngResult = MsgBox("Select the Node you wish to Move TO, THEN hit F6" & _
        vbCrLf & "  (or hit Cancel to stop the Move)", _
        vbOKCancel + vbInformation, _
        "Moving Node <" & Left$(From_Node.Text, 20) & ">")
      If lngResult = vbCancel Then
        GoTo Cleanup_Move
      Else
        Exit Sub
      End If
    End If
    On Error GoTo Cleanup_Move  'Handles moving a node to itself
    Set From_Node.Parent = To_Node
    TreeIsDirty True
Cleanup_Move:
    From_Node.Tag = StripFromTag(From_Node.Tag, "MV")
    Set_Node_Colors From_Node
    Set From_Node = Nothing
    mnuMove.Caption = "&Move"
  End If
End Sub

Private Sub mnuNodeBold_Click()
  mnuNodeBold.Checked = Not mnuNodeBold.Checked
  With TreeView
    .SelectedItem.Bold = mnuNodeBold.Checked
    .SelectedItem.ForeColor = vbBlue
    TreeIsDirty True
    '.Refresh
  End With
End Sub

Private Sub mnuNodeSorted_Click()
  mnuNodeSorted.Checked = Not mnuNodeSorted.Checked
  With TreeView
    .SelectedItem.Sorted = mnuNodeSorted.Checked
    '.Refresh
  End With
End Sub

Private Sub mnuPassword_Click()
  txtPass.Top = Me.Height - (txtPass.Height + 900)
  cmdAccept.Top = txtPass.Top
  lblPass.Top = txtPass.Top
  TreeView.Height = txtPass.Top - 100
  TreeView.SelectedItem.BackColor = vbMagenta
  TreeView.SelectedItem.ForeColor = vbBlue
  TreeView.SelectedItem.EnsureVisible
  TreeView.Enabled = False
  txtPass.Text = ""
  ITConfirmPassword = ""
  If mnuPassword.Caption = "Add &Password" Then
    lblPass.Caption = "Enter Password to Add:"
  Else
    lblPass.Caption = "Enter Password to Remove:"
  End If
  txtPass.SetFocus
End Sub

Private Sub mnuPrint_Click()
Dim CHNL As Integer
Dim File As String
Dim TopNode As Node
Dim SiblingNode As Node
  'If TreeIsDirty Then mnuSave_Click
  File = RegreadKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Desktop")
  File = File & "\InfoTree.TXT"
  ITPrintLevel = 0
  CHNL = FreeFile
  Open File For Output As #CHNL
  Set TopNode = TreeView.SelectedItem
  If Not (TopNode Is Nothing) Then
    Print #CHNL, Space(ITPrintLevel) & TopNode.Text
    PrintChildrenOf TopNode.Key, CHNL
    If (TopNode.Parent Is Nothing) Then
      'If printing highest node, print siblings as well
      Set SiblingNode = TopNode.FirstSibling
      If TopNode.Children > 1 Then
        If SiblingNode.Key = TopNode.Key Then
          Set SiblingNode = SiblingNode.Next
        End If
      End If
      If Not (SiblingNode Is Nothing) Then  'If child has siblings
        Do Until SiblingNode.Key = TopNode.Key
          'Print this Sibling
          Print #CHNL, Space(ITPrintLevel) & SiblingNode.Text
          PrintChildrenOf SiblingNode.Key, CHNL
          Set SiblingNode = SiblingNode.Next
          If (SiblingNode Is Nothing) Then Exit Do
        Loop
      End If
    End If
  End If
  Close #CHNL
  If ITFileExists(File) Then
    Shell ITGet_Environ("COMSPEC") & " /C " & Chr(34) & File & Chr(34)
  End If
  MsgBox "File <InfoTree.TXT> is on your Desktop"
End Sub

Private Sub mnuReplace_Click()
Dim AlwaysReplace As Boolean
Dim ChildNode As Node
Dim FoundFirstNode As Boolean
Dim IsDirty As Boolean
Dim ThisNode As Node
Dim PosFound As Integer
Dim lngResult As Long
Dim strQuestion As String
Dim strReplace As String
Dim YesCount As Long
  mnuFind_Click   'Find the text first.
  
  If Len(ITFind) > 0 Then
    strReplace = InputBox("Enter Replacement for <" & ITFind & ">:", "Replace In Tree", "")
    If Len(strReplace) = 0 Then Exit Sub
  
    For Each ThisNode In TreeView.Nodes
      With ThisNode
        If .BackColor = vbCyan Then
          PosFound = InStr(1, UCase(.Text), ITFind)
          Do Until PosFound = 0
            If Not AlwaysReplace Then
              ThisNode.EnsureVisible
              ThisNode.BackColor = vbDarkRed
              ThisNode.ForeColor = vbWhite
              DoEvents
              strQuestion = "Replace?" & vbCrLf & _
                "Yes    = This Line" & vbCrLf & _
                "No     = Skip to next Line" & vbCrLf & _
                "Cancel = Stop this Process"
              lngResult = MsgBox(strQuestion, vbYesNoCancel + vbQuestion, "Replace?")
              Set_Node_Colors ThisNode
              Select Case lngResult
              Case vbYes
                YesCount = YesCount + 1
                'Fall through
                If YesCount > 3 Then
                  strQuestion = "ALWAYS Replace?" & vbCrLf & _
                    "OK     = Replace ALL lines" & vbCrLf & _
                    "Cancel = No, Keep asking me"
                  lngResult = MsgBox(strQuestion, vbOKCancel + vbQuestion, "ALWAYS Replace?")
                  YesCount = -30
                  If lngResult = vbOK Then AlwaysReplace = True
                End If
              Case vbNo
                YesCount = -1 * YesCount
                Exit Do
              Case vbCancel
                Exit For
              Case Else
                MsgBox "Huh?  lngResult = " & Format(lngResult)
                Exit Do 'Should never happen
              End Select
            End If
            .Text = SetMid(.Text, strReplace, PosFound, Len(ITFind))
            PosFound = InStr(1, UCase(.Text), ITFind)
            IsDirty = True
          Loop
        End If
      End With
    Next
    If IsDirty Then TreeIsDirty True
  End If
End Sub

Private Sub mnuSave_Click()
  Save_Treeview
End Sub

Private Sub mnuShowKey_Click()
  If TreeIsDirty Then
    Save_Treeview
  End If
  mnuShowKey.Checked = Not mnuShowKey.Checked
  Load_Treeview
End Sub

Private Sub mnuSortTree_Click()
  If TreeIsDirty Then
    Save_Treeview
  End If
  mnuSortTree.Checked = Not mnuSortTree.Checked
  Load_Treeview
End Sub

Private Sub mnuUndo_Click()
Dim lngResult As Long
Dim strQuestion As String
Dim SQL As String
  strQuestion = "Are you Sure you want to restore the last saved version?"
  lngResult = MsgBox(strQuestion, vbQuestion + vbYesNo, "Confirm Undo to Last saved Tree")
  If lngResult = vbYes Then
    SQL = "DELETE * FROM InfoTree WHERE NLV=False;"
    ITdb.Execute SQL
    DoEvents
    SQL = "UPDATE InfoTree SET NLV=False;"
    ITdb.Execute SQL
    DoEvents
    ITrst.Requery
    Load_Treeview
  End If
End Sub

Private Sub TreeView_AfterLabelEdit(Cancel As Integer, NewString As String)
  With TreeView.SelectedItem
    If .Text <> NewString Then
      .Tag = AddToTag(.Tag, "ND") 'ND = Node Dirty
      Set_Node_Colors TreeView.SelectedItem
      TreeIsDirty True
    End If
  End With
End Sub

Private Sub TreeView_Click()
  With TreeView
    If Not (.SelectedItem Is Nothing) Then
      mnuNodeSorted.Checked = .SelectedItem.Sorted
      mnuNodeBold.Checked = .SelectedItem.Bold
      If InStr(1, .SelectedItem.Tag, "|PW") Then
        mnuPassword.Caption = "Remove &Password"
      Else
        mnuPassword.Caption = "Add &Password"
      End If
    End If
  End With
End Sub

Private Sub TreeView_DblClick()
  TreeView.StartLabelEdit
End Sub

Private Function CopyChildrenOf(ByVal FromNodeKey As String, _
    ByVal ToNodeKey As String, _
    ByVal ToNodeText As String)  'Recursive copy
Dim FromNode As Node
Dim ChildNode As Node
Dim SiblingNode As Node
Dim NewChildNodeID As String
Dim NewSiblingNodeID As String
Dim lngResult As Long
  Set FromNode = TreeView.Nodes(FromNodeKey)
  If (FromNode Is Nothing) Then Exit Function
  Set ChildNode = FromNode.Child
  If (ChildNode Is Nothing) Then Exit Function  'No children
  
  If FromNode.Text = ToNodeText Then
    lngResult = MsgBox( _
      "Possible Infinite Recursion:" & vbCrLf & _
        "From: " & Left$(FromNode.Parent.Text, 20) & " > " & _
          Left$(FromNode.Text, 20) & vbCrLf & _
        "To:   " & Left$(ToNodeText, 20) & vbCrLf, _
      vbQuestion + vbYesNo, _
      "Continue to Copy?")
    If lngResult = vbNo Then Exit Function
  End If
  
  'Copy this Child to the new Branch
  NewChildNodeID = NextFreeNodeID
  
  AddNode ToNodeKey, NewChildNodeID, ChildNode.Text
  CopyChildrenOf ChildNode.Key, NewChildNodeID, ToNodeText
  
  Set SiblingNode = ChildNode.FirstSibling
  If FromNode.Children > 1 Then
    If SiblingNode.Key = ChildNode.Key Then
      Set SiblingNode = SiblingNode.Next
    End If
  End If
  If Not (SiblingNode Is Nothing) Then  'If child has siblings
    Do Until SiblingNode.Key = ChildNode.Key
      'Copy this Sibling to the new Branch
      NewSiblingNodeID = NextFreeNodeID
      AddNode ToNodeKey, NewSiblingNodeID, SiblingNode.Text
      CopyChildrenOf SiblingNode.Key, NewSiblingNodeID, ToNodeText
      Set SiblingNode = SiblingNode.Next
      If (SiblingNode Is Nothing) Then Exit Do
    Loop
  End If
End Function

Private Function PrintChildrenOf(ByVal FromNodeKey As String, _
    ByVal CHNL As Integer)  'Recursive Print
Dim FromNode As Node
Dim ChildNode As Node
Dim SiblingNode As Node
  Set FromNode = TreeView.Nodes(FromNodeKey)
  If (FromNode Is Nothing) Then Exit Function
  Set ChildNode = FromNode.Child
  If (ChildNode Is Nothing) Then Exit Function  'No children
  
  'Print this Child
  ITPrintLevel = ITPrintLevel + 2
  Print #CHNL, Space(ITPrintLevel) & ChildNode.Text
  PrintChildrenOf ChildNode.Key, CHNL
  
  Set SiblingNode = ChildNode.FirstSibling
  If FromNode.Children > 1 Then
    If SiblingNode.Key = ChildNode.Key Then
      Set SiblingNode = SiblingNode.Next
    End If
  End If
  If Not (SiblingNode Is Nothing) Then  'If child has siblings
    Do Until SiblingNode.Key = ChildNode.Key
      'Print this Sibling
      Print #CHNL, Space(ITPrintLevel) & SiblingNode.Text
      PrintChildrenOf SiblingNode.Key, CHNL
      Set SiblingNode = SiblingNode.Next
      If (SiblingNode Is Nothing) Then Exit Do
    Loop
  End If
  ITPrintLevel = ITPrintLevel - 2
End Function

Private Function NextFreeNodeID() As String
    NextFreeNodeID = "K" & Format(ITNextNodeID, "00000")
    ITNextNodeID = ITNextNodeID + 1
End Function

Private Function RegreadKey(Value As String) As String
Dim Obj As Object
  On Error GoTo Exit_Function
  Set Obj = CreateObject("wscript.shell")
  RegreadKey = Obj.regread(Value)
Exit_Function:
End Function

Private Function Set_Menu_Options()
  If Len(KeyForText(HELP_NODE)) > 0 Then
    mnuAddHelp.Enabled = False
  Else
    mnuAddHelp.Enabled = True
  End If
  If Len(KeyForText(STATS_NODE)) > 0 Then
    mnuAddStats.Enabled = False
  Else
    mnuAddStats.Enabled = True
  End If
  mnuSave.Enabled = mnuCancel.Enabled  'mnuCancel functions as 'Is Dirty'.
End Function

Private Function AddToTag(ByVal iTag As String, _
  ByVal iTagDataType As String, Optional ByVal iTagData As Variant) As String
Dim itData As String
  If IsMissing(iTagData) Then
    itData = ""
  Else
    itData = iTagData
  End If
  AddToTag = iTag & "|" & iTagDataType & itData & "||ND|" ' ND = Node is Dirty
  'CAN'T DO THIS TreeIsDirty True
End Function

Private Function StripFile(ByVal FileAndPath$) ' Strips filename, leaving Path
Dim j$, X As Integer
    Do Until InStr(FileAndPath$, "\") = 0     ' Strip off the Path
        X = InStr(FileAndPath$, "\")
        j$ = j$ & Trim(Left(FileAndPath$, X))
        FileAndPath$ = Mid(FileAndPath$, X + 1, Len(FileAndPath$))
    Loop
    StripFile = j$
End Function

Private Function StripFromTag(ByVal iTag As String, _
  ByVal iTagDataType As String) As String
Dim P1 As Integer
Dim P2 As Integer
  P1 = InStr(1, iTag, "|" & iTagDataType)
  If P1 > 0 Then
    P2 = InStr(P1 + 1, iTag, "|")
    If P2 > 0 Then
      StripFromTag = ITSetMid(iTag, "", P1, (P2 - P1) + 1)
    Else
      StripFromTag = Left$(iTag, P1 - 1)
    End If
    'CAN'T DO THIS TreeIsDirty True
  Else
    StripFromTag = iTag
  End If
  End Function


Private Function SetMid(ByVal istrToModify As String, ByVal istrToInsert As String, _
  ByVal ilngStart As Long, ByVal ilngCharsToRemove As Long) As String
'Inputs:  istrToModify - The string you wish to modify.
'         istrToInsert - The string you are inserting into istrToModify.
'         ilngStart - The character In istrToModify from which istrToInsert must be inserted.
'         ilngCharsToRemove - How many characters istrToInsert must replace in
'           istrToModify, from ilngStart.
'           - Set this To 0,  istrToInsert will be inserted without over-writing any characters in istrToModify.
'           - Set this To -1, istrToInsert will over-write to the length of istrToInsert
'Returns: An appropriate merge of istrToModify and istrToInsert.
'
    Dim first As String
    Dim last As String
    If ilngCharsToRemove < 0 Then ilngCharsToRemove = Len(istrToInsert)
    first = Left(istrToModify, ilngStart - 1)
    last = Right(istrToModify, Len(istrToModify) - (ilngStart + ilngCharsToRemove) + 1)
    SetMid = first & istrToInsert & last
End Function


Private Function Set_Node_Colors(ByRef iNode As Node) As Boolean
  'Set_Node_Colors = True
  With iNode
    If Len(.Tag) < 14 Then     '"|IID123456|" and min 3 more chars makes 2 tags
      .ForeColor = vbBlack
      .BackColor = vbWhite
      Exit Function
    End If
    If InStr(1, .Tag, "|NL") > 0 Then   'Node to link to
      .ForeColor = vbGreen
      .BackColor = vbYellow
      Exit Function
    End If
    If InStr(1, .Tag, "|CP") > 0 Then   'Node to copy from
      .ForeColor = vbBlue
      .BackColor = vbYellow
      Exit Function
    End If
    If InStr(1, .Tag, "|MV") > 0 Then   'Node to Move
      .ForeColor = vbRed
      .BackColor = vbYellow
      Exit Function
    End If
    If InStr(1, .Tag, "|LNK") > 0 Then  'Existing Link
      .ForeColor = vbDarkGreen
      .BackColor = vbWhite
      Exit Function
    End If
    If InStr(1, .Tag, "|PW") > 0 Then   'Password protected field
      .ForeColor = vbMagenta
      .BackColor = vbWhite
      Exit Function
    End If
    If InStr(1, .Tag, "|ND") > 0 Then   'Node is Dirty
      .ForeColor = vbBlue
      .BackColor = vbWhite
      Exit Function
    End If
    'Should Never Happen
    .ForeColor = vbBlack
    .BackColor = vbGreen
  End With
End Function

Private Sub TreeView_Expand(ByVal Node As MSComctlLib.Node)
Dim strPW As String
  With Node
    If InStr(1, .Tag, "|PW") > 0 Then
      If DateDiff("s", ITPasswordApproved, Now()) > 2 Then
        Node.Expanded = False
        DoEvents
        strPW = InputBox("Enter Password for Node <" & Left(Node.Text, 20) & ">", _
          "Password Required!", "")
        If strPW <> "" Then
          If InStr(1, UCase(Node.Tag), "|PW" & UCase(strPW)) = 0 Then
            Beep
          Else
            ITPasswordApproved = Now()
            Node.Expanded = True
          End If
        End If
      Else
        ITPasswordApproved = 0
      End If
    End If
  End With
End Sub

