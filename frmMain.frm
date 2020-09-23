VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Demo Info Tree"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescription 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmMain.frx":0000
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdShowTree 
      Caption         =   "Show Info Tree (1 line of code behind this button)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdShowTree_Click()
  'The following line activates the InfoTree
  frmInfoTree.Show
End Sub

Private Sub Form_Load()
  With txtDescription
    .Text = "Instructions to add InfoTree to your Application:" & vbCrLf
    .Text = .Text & "In addition to frmInfoTree.frm, you'll need to add "
    .Text = .Text & "the following components to your project: " & vbCrLf
    .Text = .Text & "1) Microsoft Windows Common Controls 6.0 (SP3)" & vbCrLf
    .Text = .Text & "   (MSCOMCTL.OCX)" & vbCrLf
    .Text = .Text & "2) Microsoft DAO 3.5 Object Library (or later)" & vbCrLf
    .Text = .Text & "   (DAO350.DLL)" & vbCrLf
    .Text = .Text & "3) TREE2.ICO - Icon for the frmInfoTree form." & vbCrLf
    .Text = .Text & vbCrLf
    .Text = .Text & "Help using the InfoTree will be displayed IN the InfoTree." & vbCrLf
    .Text = .Text & "P.S. - the database (InfoTree.ITR) file is actually a " & _
            "standard Jet Database (.mdb) and can be viewed using MS-Access by renaming " & _
            "the .ITR file to .MDB." & vbCrLf
    .Text = .Text & "Also, All passwords in the test data are 'x' (without the quotes)."
  End With

End Sub

