VERSION 5.00
Begin VB.Form FrmID3 
   Caption         =   "ID3 Editor"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtComments 
      Height          =   285
      Left            =   960
      MaxLength       =   30
      TabIndex        =   13
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtGenre 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   960
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtAlbum 
      Height          =   285
      Left            =   960
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   960
      MaxLength       =   30
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtSong 
      Height          =   285
      Left            =   960
      MaxLength       =   30
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtFilename 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Comments:"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Genre: "
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Year:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Album:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Artist:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Song:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mvarFilename As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sInfo As Id3
    
    sInfo.Album = Left(txtAlbum.Text, 30)
    sInfo.Artist = Left(txtArtist.Text, 30)
    sInfo.Title = Left(txtSong.Text, 30)
    sInfo.sYear = Left(txtYear.Text, 4)
    sInfo.Comments = Left(txtComments.Text, 30)
    'sInfo.Genre = txtGenre.Text
    
    Id3Module.SaveId3 mvarFilename, sInfo
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sInfo As New clsMP3Info
    
    sInfo.Filename = mvarFilename
    
    txtFilename.Text = mvarFilename
    txtSong.Text = sInfo.Title
    txtArtist.Text = sInfo.Artist
    txtAlbum.Text = sInfo.Album
    txtYear.Text = sInfo.Year
    txtGenre.Text = sInfo.Genre
    txtComments.Text = sInfo.Comment
    
    Set sInfo = Nothing
End Sub

Public Property Let Filename(ByVal sData As String)
    mvarFilename = sData
End Property

Public Property Get Filename() As String
    Filename = mvarFilename
End Property
