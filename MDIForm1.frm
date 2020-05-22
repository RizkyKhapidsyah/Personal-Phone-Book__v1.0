VERSION 5.00
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Phone Book v1.0"
   ClientHeight    =   4650
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuContacts 
      Caption         =   "&Contacts"
      Begin VB.Menu mnuShowContactList 
         Caption         =   "Show / Hide Contact &List"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DB As Database
Public ContactTable As Recordset

Private Sub MDIForm_Load()
    mnuShowContactList_Click
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Public Sub mnuShowContactList_Click()
    If Not frmContList.Visible Then
        Load frmContList
        frmContList.Top = 60
        frmContList.Left = 60
        frmContList.Show
    Else
        Unload frmContList
    End If
End Sub
