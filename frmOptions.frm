VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Player Options"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox cbSavePlaylist 
      Caption         =   "Save Playlist on Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.CheckBox cbDelete 
      Caption         =   "Confirm Song Deletion"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CheckBox cbScroll 
      Caption         =   "Scroll Title"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CheckBox cbSave 
      Caption         =   "Save Settings on Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Player Title"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbScroll_Click()
    If cbScroll.Value = vbChecked Then
        frmMain.bScrollTitle = True
    Else
        frmMain.bScrollTitle = False
    End If
End Sub

Private Sub Form_Load()
    Text1.Text = frmMain.sPlayerName
    
    If frmMain.bScrollTitle Then
        cbScroll.Value = vbChecked
    End If
    
    If frmMain.bSaveSettings Then
        cbSave.Value = vbChecked
    End If
    
    If frmMain.bSavePlaylist Then
        cbSavePlaylist.Value = vbChecked
    End If
    
    If frmMain.bConfirmDelete Then
        cbDelete.Value = vbChecked
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.bSavePlaylist = Me.cbSavePlaylist
    frmMain.bSaveSettings = Me.cbSave
    frmMain.bConfirmDelete = Me.cbDelete
    frmMain.bScrollTitle = Me.cbScroll
    frmMain.sPlayerName = Text1.Text
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Index As Integer
    
    Index = CInt(frmMain.lstPlayList1.SelectedItem)

    frmMain.sPlayerName = Text1.Text & Chr(KeyAscii)
    frmMain.sFormTitle = frmMain.sPlayerName & " - [" & frmMain.lstPlayList1.ListItems(Index).ListSubItems(1).Text & "-" & frmMain.lstPlayList1.ListItems(Index).ListSubItems(2).Text & "]  "
    frmMain.Caption = frmMain.sFormTitle
End Sub
