VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BA9BDE51-F9CF-11D3-B7A6-00A0CC290A67}#1.0#0"; "HCSlider.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MP3 Player"
   ClientHeight    =   6150
   ClientLeft      =   105
   ClientTop       =   -180
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin HCSlider.HCSSlider HCSSlider1 
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   2280
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "Reloads settings and bitmaps (so you don't have to exit if you change stuff)"
      Top             =   5640
      Width           =   615
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   0
      ScaleHeight     =   3000
      ScaleWidth      =   3750
      TabIndex        =   15
      Top             =   2520
      Width           =   3750
      Begin MSComctlLib.ListView lstPlayList1 
         Height          =   2655
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   0
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Number"
            Object.Width           =   661
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SongTitle"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Artist"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   1800
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Clear the Playlist"
      Top             =   5640
      Width           =   495
   End
   Begin VB.CheckBox cbShuffle 
      BackColor       =   &H00000000&
      Caption         =   "Shuffle"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CheckBox cbRepeat 
      BackColor       =   &H00000000&
      Caption         =   "Repeat"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   2
      ToolTipText     =   "Save the playlist to a file"
      Top             =   5640
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.ListBox lstFilenames 
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   3750
      TabIndex        =   10
      Top             =   1440
      Width           =   3750
      Begin MSComctlLib.Slider Slider2 
         Height          =   135
         Left            =   2760
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   238
         _Version        =   393216
         Min             =   -3000
         Max             =   3000
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   135
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   238
         _Version        =   393216
         Min             =   -3000
         Max             =   0
         TickStyle       =   3
      End
      Begin VB.Image imgNext 
         Height          =   375
         Left            =   2160
         ToolTipText     =   "Next Track"
         Top             =   240
         Width           =   375
      End
      Begin VB.Image imgPrev 
         Height          =   375
         Left            =   240
         ToolTipText     =   "Previous Track"
         Top             =   240
         Width           =   375
      End
      Begin VB.Image imgStop 
         Height          =   375
         Left            =   1680
         ToolTipText     =   "Stop"
         Top             =   240
         Width           =   375
      End
      Begin VB.Image imgPause 
         Height          =   375
         Left            =   1200
         ToolTipText     =   "Pause"
         Top             =   240
         Width           =   375
      End
      Begin VB.Image imgPlay 
         Height          =   375
         Left            =   720
         ToolTipText     =   "Play"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   0
      ScaleHeight     =   1500
      ScaleWidth      =   3750
      TabIndex        =   7
      Top             =   0
      Width           =   3750
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   1
      ToolTipText     =   "Add files to the playlist"
      Top             =   5640
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   5520
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label cmdShowPlaylist 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Show Playlist"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   3735
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   135
      Left            =   720
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -380
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuSort 
      Caption         =   "Sort"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHidePlaylist 
         Caption         =   "&Hide Playlist"
      End
      Begin VB.Menu mnuBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortBy 
         Caption         =   "Sort &By"
         Begin VB.Menu mnuArtist 
            Caption         =   "&Artist/Songtitle"
         End
         Begin VB.Menu mnuTitle 
            Caption         =   "&Songtitle"
         End
         Begin VB.Menu mnuFilename 
            Caption         =   "&Filename"
         End
         Begin VB.Menu mnuBar4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRandom 
            Caption         =   "&Randomize"
         End
         Begin VB.Menu mnuReverse 
            Caption         =   "R&everse Order"
         End
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInfo 
         Caption         =   "&File Info"
      End
      Begin VB.Menu mnuPref 
         Caption         =   "&Preferences"
         Begin VB.Menu mnuFont 
            Caption         =   "&Font"
         End
         Begin VB.Menu mnuFontColor 
            Caption         =   "Font &Color"
         End
         Begin VB.Menu mnuBackColor 
            Caption         =   "&Back Color"
         End
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Form"
      Visible         =   0   'False
      Begin VB.Menu mnuHidePlaylist2 
         Caption         =   "&Hide Playlist"
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "&Compact View"
      End
      Begin VB.Menu mnuPref2 
         Caption         =   "&Preferences"
         Begin VB.Menu mnuFont2 
            Caption         =   "&Font"
         End
         Begin VB.Menu mnuFontColor2 
            Caption         =   "Font &Color"
         End
         Begin VB.Menu mnuBackColor2 
            Caption         =   "&Back Color"
         End
         Begin VB.Menu mnuOptions 
            Caption         =   "&Options"
         End
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuControls 
      Caption         =   "Controls"
      Visible         =   0   'False
      Begin VB.Menu mnuFullView 
         Caption         =   "&Full View"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const gMISSINGREGDATA = "$$EMPTY$$"
Const gEMPTYSTRING = ""
Const gKEYNAME = "SoftWare\MP3Player"   'reg settings

Public bSaveSettings As Boolean
Public bSavePlaylist As Boolean
Public bScrollTitle As Boolean
Public bConfirmDelete As Boolean
Public sPlayerName As String
Public sFormTitle As String

Enum enFontObject
    Playlist = 2
    Other = 4
End Enum


Dim Mp3Info As New clsMP3Info
Dim bMove As Boolean
Dim iOldX As Integer
Dim iOldY As Integer
Dim iCurrentIndex As Integer
Dim iPrevIndex As Integer
Dim sImageDir As String
Dim bStopPressed As Boolean
Dim sStartDir As String
Dim gFontColor1 As Long
Dim gFontColor2 As Long

Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub cmdClear_Click()
'Purpose:  Clear the playlists
    lstFilenames.Clear
    lstPlayList1.ListItems.Clear
End Sub

'For saving playlists
Private Sub cmdSave_Click()
    CommonDialog1.Filename = ""
    CommonDialog1.Filter = "Playlist (*.mpl)|*.mpl"
    CommonDialog1.ShowSave
    If CommonDialog1.Filename <> "" Then
        SavePlaylist CommonDialog1.Filename
    End If
End Sub

'Reshows the playlist when it's hidden
Private Sub cmdShowPlaylist_Click()
    Call mnuHidePlaylist_Click
End Sub

'changes color of the 'Show Playlist' text when it's clicked
Private Sub cmdShowPlaylist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdShowPlaylist.ForeColor = lstPlayList1.ForeColor
End Sub

'changes color of 'Show Playlist' text back when the mouse button is released
Private Sub cmdShowPlaylist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdShowPlaylist.ForeColor = Label1.ForeColor
End Sub


Private Sub cmdOpen_Click()
    CommonDialog1.Filename = ""
    CommonDialog1.DefaultExt = ".mpl"
    CommonDialog1.Filter = "MP3 Audio (*.mp3)|*.mp3|Wave Files (*.wav)|*.wav|MIDI Files (*.mid)|*.mid|PlayList (*.mpl)|*.mpl|All Files (*.*)|*.*"
    CommonDialog1.Flags = cdlOFNOverwritePrompt
    CommonDialog1.ShowOpen
    'Check to see if it's a playlist or just a song and import accordingly
    If CommonDialog1.Filename <> "" And Right(CommonDialog1.Filename, 3) <> "mpl" Then
        If ParseFiles(CommonDialog1.Filename) Then
            lstFilenames.AddItem CommonDialog1.Filename
        End If
    ElseIf CommonDialog1.Filename <> "" And Right(CommonDialog1.Filename, 3) = "mpl" Then
        LoadPlayList CommonDialog1.Filename
    End If
End Sub

'Reloads the bitmaps, and color/font settings
Private Sub cmdRefresh_Click()
    LoadINIFile
    LoadImages
    SetPictures
End Sub

Private Sub Form_Load()
    Dim iFilenum As Integer
    Dim sTemp As String
    Dim iIndex As Integer
    
    'Set initial properties in the MediaPlayer control
    MediaPlayer1.AutoStart = False
    MediaPlayer1.AutoRewind = True
    MediaPlayer1.ShowAudioControls = True
    MediaPlayer1.Volume = 0
    'Set default values for program variables
    bStopPressed = True
    iCurrentIndex = 1
    gFontColor1 = -1
    gFontColor2 = -1
    
    'Sets the directory where the MP3 player is located (used for saving, loading, etc.)
    sStartDir = RegistryQuery("Path", "Enter the FULL path where MP3 Player is installed")
    
    'Format the sStartDir so I know it has a '\' at the end
    If Right(sStartDir, 1) <> "\" Then
        sStartDir = sStartDir & "\"
    End If
    
    'Sets the directory where the images are kept
    sImageDir = sStartDir & "images\"
    
    'Loads preferences from and ini file
    LoadINIFile
    
    'Loads the images from the image directory into an imagelist control
    LoadImages
    'Sets the images from the imagesList into the appropriate places on the form
    SetPictures
    
    'Checks for the existance of the playlist that is automatically saved on close
    If Dir(sStartDir & "\playlist.mpl") <> "" Then
        'loads the saved playlist's songs
        LoadPlayList sStartDir & "\playlist.mpl"
    End If
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bMove = True 'For moving the form without having to drag the title bar
    ElseIf Button = vbRightButton Then
        frmMain.PopupMenu mnuForm 'shows popup menu
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bMove Then
        'moves the form
        frmMain.Move frmMain.Left + (X - iOldX), frmMain.Top + (Y - iOldY)
    Else
        iOldX = X
        iOldY = Y
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        bMove = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Me.bSavePlaylist Then
        'Save songs to a file to be loaded when re-opened
        SavePlaylist sStartDir & "\playlist.mpl"
    Else
        'If they choose not to have the playlist automatically saved, and an old
        'playlist exists, it is deleted
        If Dir(sStartDir & "\playlist.mpl") <> "" Then
            Kill sStartDir & "\playlist.mpl"
        End If
    End If
    
    'Save the settings if the option is checked
    'Otherwise, delete the ini file...it is recreated with default values next
    'time the player is opened.
    If Me.bSaveSettings Then
        SaveINIFile
    Else
        If Dir(sImageDir & "skin.ini") <> "" Then
            Kill sImageDir & "skin.ini"
        End If
    End If
    
    'Destroy the ID3 parsing object
    Set Mp3Info = Nothing
End Sub

Private Sub HCSSlider1_Moved()
    MediaPlayer1.CurrentPosition = (HCSSlider1.Percent / 100) * MediaPlayer1.Duration
End Sub

Private Sub imgNext_Click()
'Purpose:  Plays the next song, taking shuffle/repeat into account,
'          as well as the position in the playlist (if it's at the end, it moves
'          to the beginning)
    Dim iIndex As Integer
    
    'Check to see if shuffle is on
    If cbShuffle.Value = vbChecked Then
        Randomize 'Initialize random number generator with system time
        iIndex = GetRandomIndex(lstFilenames.ListCount - 1) 'Get random index
        While iIndex = iCurrentIndex 'Make sure random index isn't the current song
            iIndex = GetRandomIndex(lstFilenames.ListCount - 1)
        Wend
        PlaySong iIndex 'Play randomly generated song
    Else
        'Check to see if it's at the end of the list
        If iCurrentIndex = lstFilenames.ListCount Then
            'If it's set to repeat, play the first song,
            'otherwise, just load it and stop
            If cbRepeat.Value = vbChecked Then
                PlaySong 1
            Else
                bStopPressed = True
                LoadSong 1
            End If
        Else
            'If it's not at the end, simply play the next song
            PlaySong iCurrentIndex + 1
        End If
    End If
End Sub

Private Sub imgNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgNext.Picture = ImageList1.ListImages(6).Picture
End Sub

Private Sub imgNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgNext.Picture = ImageList1.ListImages(5).Picture
End Sub

Private Sub imgPause_Click()
    If MediaPlayer1.PlayState = mpPlaying Then
        MediaPlayer1.Pause
    Else
        MediaPlayer1.Play
    End If
End Sub

Private Sub imgPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPause.Picture = ImageList1.ListImages(9).Picture
End Sub

Private Sub imgPlay_Click()
    PlaySong lstPlayList1.SelectedItem
End Sub

Private Sub imgPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPlay.Picture = ImageList1.ListImages(12).Picture
End Sub

Private Sub imgPrev_Click()
    Dim iIndex As Integer
        
    If cbShuffle.Value = vbChecked Then
        Randomize
        iIndex = GetRandomIndex(lstFilenames.ListCount - 1)
        PlaySong iIndex
    Else
        If iCurrentIndex = 1 Then
            PlaySong lstPlayList1.ListItems.Count
        Else
            PlaySong iCurrentIndex - 1
        End If
    End If
End Sub

Private Sub imgPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrev.Picture = ImageList1.ListImages(14).Picture
End Sub

Private Sub imgPrev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgPrev.Picture = ImageList1.ListImages(13).Picture
End Sub

Private Sub imgStop_Click()
    MediaPlayer1.Stop
    MediaPlayer1.CurrentPosition = 0
    bStopPressed = True 'A hack that fixes a bug where the next song wouldn't play
                        'because the total time had a small decimal after it.  Mediaplayer
                        'would stop before the program realized it was at the end of the song
End Sub

Private Sub imgStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set imgStop.Picture = ImageList1.ListImages(18).Picture
End Sub

Private Sub lstPlayList1_DblClick()
    PlaySong lstPlayList1.SelectedItem
End Sub

Private Sub lstPlayList1_KeyDown(KeyCode As Integer, Shift As Integer)
'Purpose:  Removes a song from the playlist
    Dim iIndex As Integer
    Dim iCnt As Integer

    If KeyCode = vbKeyDelete Then
        DeleteFile CInt(lstPlayList1.SelectedItem.Text)
    End If
End Sub

Private Sub lstPlayList1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Set lstPlayList1.SelectedItem = lstPlayList1.HitTest(X, Y)
        frmMain.PopupMenu mnuSort
    End If
End Sub

Private Sub lstPlayList1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iCnt As Integer
    
    For iCnt = 1 To Data.Files.Count
        If InStr(1, Data.Files(iCnt), ".") <> 0 Then
            If ParseFiles(Data.Files(iCnt)) Then
                lstFilenames.AddItem Data.Files(iCnt)
            End If
        Else
            AddDirectory Data.Files(iCnt)
        End If
    Next iCnt
End Sub

Private Sub MediaPlayer1_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
'Purpose:  Makes sure the bitmaps for the play controls match the state of the mediaplayer
    Select Case MediaPlayer1.PlayState
    
    Case mpClosed
        Set imgStop.Picture = ImageList1.ListImages(17).Picture
        Set imgPlay.Picture = ImageList1.ListImages(10).Picture
        Set imgPause.Picture = ImageList1.ListImages(7).Picture
    Case mpPaused
        Set imgStop.Picture = ImageList1.ListImages(16).Picture
        Set imgPlay.Picture = ImageList1.ListImages(11).Picture
        Set imgPause.Picture = ImageList1.ListImages(8).Picture
    Case mpPlaying
        Set imgStop.Picture = ImageList1.ListImages(16).Picture
        Set imgPlay.Picture = ImageList1.ListImages(11).Picture
        Set imgPause.Picture = ImageList1.ListImages(7).Picture
    Case mpStopped
        Set imgStop.Picture = ImageList1.ListImages(17).Picture
        Set imgPlay.Picture = ImageList1.ListImages(10).Picture
        Set imgPause.Picture = ImageList1.ListImages(7).Picture
    End Select
End Sub

Private Sub mnuArtist_Click()
    Dim sInfo As New clsMP3Info
    Dim iFoundAt As Integer
    Dim iCnt As Integer
    Dim iPos As Integer
    Dim sArtist As String
    Dim sTitle As String
    Dim sTemp As String
    
    For iCnt = 1 To lstPlayList1.ListItems.Count
        sArtist = lstPlayList1.ListItems(iCnt).ListSubItems(2).Text
        sTitle = lstPlayList1.ListItems(iCnt).ListSubItems(1).Text
        iFoundAt = iCnt - 1
        sTemp = lstFilenames.List(iCnt - 1)
        For iPos = iCnt To lstPlayList1.ListItems.Count
            If sArtist > lstPlayList1.ListItems(iPos).ListSubItems(2).Text Then
                sTemp = lstFilenames.List(iPos - 1)
                sArtist = lstPlayList1.ListItems(iPos).ListSubItems(2).Text
                sTitle = lstPlayList1.ListItems(iPos).ListSubItems(1).Text
                iFoundAt = iPos - 1
            ElseIf sArtist = lstPlayList1.ListItems(iPos).ListSubItems(2).Text Then
                If sTitle > lstPlayList1.ListItems(iPos).ListSubItems(1).Text Then
                    sTemp = lstFilenames.List(iPos - 1)
                    sTitle = lstPlayList1.ListItems(iPos).ListSubItems(1).Text
                    iFoundAt = iPos - 1
                End If
            End If
        Next iPos
        lstFilenames.List(iFoundAt) = lstFilenames.List(iCnt - 1)
        lstFilenames.List(iCnt - 1) = sTemp
        lstPlayList1.ListItems(iFoundAt + 1).ListSubItems(1).Text = lstPlayList1.ListItems(iCnt).ListSubItems(1).Text
        lstPlayList1.ListItems(iFoundAt + 1).ListSubItems(2).Text = lstPlayList1.ListItems(iCnt).ListSubItems(2).Text
        lstPlayList1.ListItems(iCnt).ListSubItems(1).Text = sTitle
        lstPlayList1.ListItems(iCnt).ListSubItems(2).Text = sArtist
    Next iCnt
    
'    lstPlayList1.ListItems.Clear
'    For iCnt = 0 To lstFilenames.ListCount - 1
'        ParseFiles lstFilenames.List(iCnt)
'    Next iCnt
End Sub

Private Sub mnuBackColor_Click()
    CommonDialog1.Color = lstPlayList1.BackColor
    CommonDialog1.ShowColor
    lstPlayList1.BackColor = CommonDialog1.Color
End Sub

Private Sub mnuBackColor2_Click()
    CommonDialog1.Color = frmMain.BackColor
    CommonDialog1.ShowColor
    frmMain.BackColor = CommonDialog1.Color
    cbRepeat.BackColor = CommonDialog1.Color
    cbShuffle.BackColor = CommonDialog1.Color
End Sub

Private Sub mnuClear_Click()
    lstFilenames.Clear
    lstPlayList1.ListItems.Clear
End Sub

Private Sub mnuCompact_Click()
    
    If mnuCompact.Checked = False Then
        Picture3.Top = 0
        frmMain.Height = Picture3.Height + 300
        mnuCompact.Checked = True
    Else
        Picture3.Top = Picture2.Height
        frmMain.Height = cbShuffle.Top + cbShuffle.Height + 500
        mnuCompact.Checked = False
    End If
End Sub

Private Sub mnuDelete_Click()
    DeleteFile CInt(lstPlayList1.SelectedItem.Text)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFileInfo_Click()
    Dim sInfo As New clsMP3Info
        
    If MediaPlayer1.Filename <> lstFilenames.List(CInt(lstPlayList1.SelectedItem) - 1) Then
        FrmID3.Filename = lstFilenames.List(CInt(lstPlayList1.SelectedItem) - 1)
        FrmID3.Show vbModal
    
        sInfo.Filename = lstFilenames.List(CInt(lstPlayList1.SelectedItem) - 1)
    
        lstPlayList1.SelectedItem.ListSubItems(1).Text = sInfo.Title
        lstPlayList1.SelectedItem.ListSubItems(2).Text = sInfo.Artist
    Else
        MsgBox "Unable to write ID3 tag." & vbCrLf & "The file might be in use.", vbCritical, "Dumbass"
    End If
    
    Set sInfo = Nothing
End Sub

Private Sub mnuFilename_Click()
    Dim iFoundAt As Integer
    Dim iCnt As Integer
    Dim iPos As Integer
    Dim sTemp As String
    
    For iCnt = 0 To lstFilenames.ListCount - 1
        sTemp = lstFilenames.List(iCnt)
        iFoundAt = iCnt
        For iPos = iCnt To lstFilenames.ListCount - 1
            If sTemp > lstFilenames.List(iPos) Then
                sTemp = lstFilenames.List(iPos)
                iFoundAt = iPos
            End If
        Next iPos
        lstFilenames.List(iFoundAt) = lstFilenames.List(iCnt)
        lstFilenames.List(iCnt) = sTemp
    Next iCnt
    
    lstPlayList1.ListItems.Clear
    For iCnt = 0 To lstFilenames.ListCount - 1
        ParseFiles lstFilenames.List(iCnt)
    Next iCnt
End Sub

Private Sub mnuFont_Click()
    CommonDialog1.Flags = cdlCFBoth
    
    CommonDialog1.FontName = lstPlayList1.Font.Name
    CommonDialog1.FontSize = lstPlayList1.Font.Size
    CommonDialog1.ShowFont
    
    If CommonDialog1.FontSize > 8 Then
        CommonDialog1.FontSize = 8
    End If
    
    ChangeFont Playlist, CommonDialog1.FontName, CommonDialog1.FontSize, CommonDialog1.FontBold, CommonDialog1.FontItalic, gFontColor1
End Sub

Private Sub mnuFont2_Click()
    CommonDialog1.Flags = cdlCFBoth
    
    CommonDialog1.FontName = frmMain.cbRepeat.Font.Name
    CommonDialog1.FontSize = frmMain.cbRepeat.Font.Size
    CommonDialog1.ShowFont
    
    If CommonDialog1.FontSize > 8 Then
        CommonDialog1.FontSize = 8
    End If
    
    ChangeFont Other, CommonDialog1.FontName, CommonDialog1.FontSize, CommonDialog1.FontBold, CommonDialog1.FontItalic, gFontColor2
End Sub

Private Sub mnuFontColor_Click()
    CommonDialog1.Color = lstPlayList1.ForeColor
    CommonDialog1.ShowColor
    gFontColor1 = CommonDialog1.Color
    
    ChangeFont Playlist, , , , , gFontColor1
End Sub

Private Sub mnuFontColor2_Click()
    CommonDialog1.Color = frmMain.cbRepeat.ForeColor
    CommonDialog1.ShowColor
    gFontColor2 = CommonDialog1.Color
    
    ChangeFont Other, , , , , gFontColor2
End Sub

Private Sub mnuFullView_Click()
    mnuCompact_Click
End Sub

Private Sub mnuHidePlaylist_Click()
    Dim i As Integer
    Dim iTotal As Integer
    If mnuHidePlaylist.Checked = False Then
        Picture4.Visible = False
        For i = 1 To 50
            cmdOpen.Top = cmdOpen.Top - (Picture4.Height / 54)
            cmdRefresh.Top = cmdRefresh.Top - (Picture4.Height / 54)
            cmdClear.Top = cmdClear.Top - (Picture4.Height / 54)
            cmdSave.Top = cmdSave.Top - (Picture4.Height / 54)
            cbRepeat.Top = cbRepeat.Top - (Picture4.Height / 54)
            cbShuffle.Top = cbShuffle.Top - (Picture4.Height / 54)
            frmMain.Height = frmMain.Height - (Picture4.Height / 54)
            frmMain.Refresh
            DoEvents
        Next i
        mnuHidePlaylist.Checked = True
        mnuHidePlaylist2.Checked = True
        cmdShowPlaylist.Visible = True
    Else
        cmdShowPlaylist.Visible = False
        For i = 1 To 50
            cmdOpen.Top = cmdOpen.Top + (Picture4.Height / 54)
            cmdRefresh.Top = cmdRefresh.Top + (Picture4.Height / 54)
            cmdClear.Top = cmdClear.Top + (Picture4.Height / 54)
            cmdSave.Top = cmdSave.Top + (Picture4.Height / 54)
            cbRepeat.Top = cbRepeat.Top + (Picture4.Height / 54)
            cbShuffle.Top = cbShuffle.Top + (Picture4.Height / 54)
            frmMain.Height = frmMain.Height + (Picture4.Height / 54)
            frmMain.Refresh
            DoEvents
        Next i
        Picture4.Visible = True
        mnuHidePlaylist.Checked = False
        mnuHidePlaylist2.Checked = False
    End If
End Sub

Private Sub mnuHidePlaylist2_Click()
    Dim i As Integer
    
    If mnuHidePlaylist2.Checked = False Then
        Picture4.Visible = False
        For i = 1 To 50
            cmdOpen.Top = cmdOpen.Top - (Picture4.Height / 54)
            cmdRefresh.Top = cmdRefresh.Top - (Picture4.Height / 54)
            cmdClear.Top = cmdClear.Top - (Picture4.Height / 54)
            cmdSave.Top = cmdSave.Top - (Picture4.Height / 54)
            cbRepeat.Top = cbRepeat.Top - (Picture4.Height / 54)
            cbShuffle.Top = cbShuffle.Top - (Picture4.Height / 54)
            frmMain.Height = frmMain.Height - (Picture4.Height / 54)
            frmMain.Refresh
            DoEvents
        Next i
        mnuHidePlaylist.Checked = True
        mnuHidePlaylist2.Checked = True
        cmdShowPlaylist.Visible = True
    Else
        cmdShowPlaylist.Visible = False
        For i = 1 To 50
            cmdOpen.Top = cmdOpen.Top + (Picture4.Height / 54)
            cmdRefresh.Top = cmdRefresh.Top + (Picture4.Height / 54)
            cmdClear.Top = cmdClear.Top + (Picture4.Height / 54)
            cmdSave.Top = cmdSave.Top + (Picture4.Height / 54)
            cbRepeat.Top = cbRepeat.Top + (Picture4.Height / 54)
            cbShuffle.Top = cbShuffle.Top + (Picture4.Height / 54)
            frmMain.Height = frmMain.Height + (Picture4.Height / 54)
            frmMain.Refresh
            DoEvents
        Next i
        Picture4.Visible = True
        mnuHidePlaylist.Checked = False
        mnuHidePlaylist2.Checked = False
    End If
End Sub

Private Sub mnuOpen_Click()
    cmdOpen_Click
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub

Private Sub mnuRandom_Click()
    Dim iCnt As Integer
    Dim sTemp As String
    Dim iNewIndex As Integer
    
    For iCnt = 0 To lstFilenames.ListCount - 1
        iNewIndex = GetRandomIndex(lstFilenames.ListCount - 1)
        sTemp = lstFilenames.List(iNewIndex)
        lstFilenames.List(iNewIndex) = lstFilenames.List(iCnt)
        lstFilenames.List(iCnt) = sTemp
    Next iCnt
    
    lstPlayList1.ListItems.Clear
    For iCnt = 0 To lstFilenames.ListCount - 1
        ParseFiles lstFilenames.List(iCnt)
    Next iCnt
End Sub

Private Sub mnuReverse_Click()
    Dim iTop As Integer
    Dim iBottom As Integer
    Dim sTemp As String
    
    iTop = 0
    iBottom = lstFilenames.ListCount - 1
    
    While iTop <= iBottom
        
        sTemp = lstFilenames.List(iBottom)
        lstFilenames.List(iBottom) = lstFilenames.List(iTop)
        lstFilenames.List(iTop) = sTemp
        
        iTop = iTop + 1
        iBottom = iBottom - 1
    Wend
    
    lstPlayList1.ListItems.Clear
    For iTop = 0 To lstFilenames.ListCount - 1
        ParseFiles lstFilenames.List(iTop)
    Next iTop
End Sub

Private Sub mnuSave_Click()
    cmdSave_Click
End Sub

Private Sub mnuTitle_Click()
    Dim iFoundAt As Integer
    Dim iCnt As Integer
    Dim iPos As Integer
    Dim sArtist As String
    Dim sTitle As String
    Dim sTemp As String
    
    For iCnt = 1 To lstPlayList1.ListItems.Count
        sArtist = lstPlayList1.ListItems(iCnt).ListSubItems(2).Text
        sTitle = lstPlayList1.ListItems(iCnt).ListSubItems(1).Text
        iFoundAt = iCnt - 1
        sTemp = lstFilenames.List(iCnt - 1)
        For iPos = iCnt To lstPlayList1.ListItems.Count
            If sTitle > lstPlayList1.ListItems(iPos).ListSubItems(1).Text Then
                sTemp = lstFilenames.List(iPos - 1)
                sArtist = lstPlayList1.ListItems(iPos).ListSubItems(2).Text
                sTitle = lstPlayList1.ListItems(iPos).ListSubItems(1).Text
                iFoundAt = iPos - 1
            End If
        Next iPos
        lstFilenames.List(iFoundAt) = lstFilenames.List(iCnt - 1)
        lstFilenames.List(iCnt - 1) = sTemp
        lstPlayList1.ListItems(iFoundAt + 1).ListSubItems(1).Text = lstPlayList1.ListItems(iCnt).ListSubItems(1).Text
        lstPlayList1.ListItems(iFoundAt + 1).ListSubItems(2).Text = lstPlayList1.ListItems(iCnt).ListSubItems(2).Text
        lstPlayList1.ListItems(iCnt).ListSubItems(1).Text = sTitle
        lstPlayList1.ListItems(iCnt).ListSubItems(2).Text = sArtist
    Next iCnt
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If mnuCompact.Checked = True Then
            frmMain.PopupMenu mnuControls
        End If
    End If
End Sub

Private Sub Slider1_Scroll()
    MediaPlayer1.Volume = Slider1.Value
End Sub

Private Sub Slider2_Scroll()
    MediaPlayer1.Balance = Slider2.Value
End Sub

Private Sub Timer1_Timer()
'Purpose:  Updates labels and keeps songs playing continuously
    Dim sStatus As String
    
    Select Case MediaPlayer1.PlayState
    
    Case mpClosed
        sStatus = "No file selected        "
    Case mpPaused
        sStatus = "Paused                  "
    Case mpPlaying
        sStatus = "Playing                 "
        HCSSlider1.Value = MediaPlayer1.CurrentPosition
    Case mpStopped
        If Not bStopPressed Then
            imgNext_Click
        End If
        sStatus = "Stopped                 "
    End Select
    
    If frmMain.bScrollTitle Then
        If sFormTitle <> "" Then
            sFormTitle = Right(sFormTitle, Len(sFormTitle) - 1) & Left(sFormTitle, 1)
        End If
        
        frmMain.Caption = sFormTitle
    End If
    
    If MediaPlayer1.PlayState <> mpClosed And MediaPlayer1.PlayState <> mpStopped Then
        Label2.Caption = sStatus & SecondsToTime(MediaPlayer1.CurrentPosition) & "/" & SecondsToTime(MediaPlayer1.Duration)
    Else
        Label2.Caption = sStatus & "0:00/0:00"
    End If
    
    If Round(MediaPlayer1.CurrentPosition, 0) >= Fix(MediaPlayer1.Duration) Then
        imgNext_Click
    End If
    
End Sub

Private Sub DeleteFile(iPlaylistIndex As Integer)
    Dim iIndex As Integer
    Dim iCnt As Integer
    Dim iResult As Integer
    
    If bConfirmDelete Then
        iResult = MsgBox("Are you sure you want to remove this song from your playlist?" & vbCrLf & "The song will still be on your computer at:" & vbCrLf & lstFilenames.List(iPlaylistIndex - 1), vbYesNo + vbQuestion, "Are you sure?")
    Else
        iResult = vbYes
    End If
    
    If iResult = vbYes Then
        lstFilenames.RemoveItem iPlaylistIndex - 1
        lstPlayList1.ListItems.Remove iPlaylistIndex
            
        For iCnt = 1 To lstPlayList1.ListItems.Count
            lstPlayList1.ListItems(iCnt).Text = iCnt
        Next iCnt
    End If
End Sub

Private Sub ChangeFont(Object As enFontObject, Optional sFontName As String, Optional sFontSize As Integer, Optional sFontBold As Boolean, Optional sFontItalic As Boolean, Optional sFontColor As Long)
    
    If Object = Playlist Then
        If sFontName <> "" Then
            lstPlayList1.Font.Name = sFontName
        End If
        
        If sFontSize > 0 Then
            lstPlayList1.Font.Size = sFontSize
        End If
        
        If Not IsMissing(sFontBold) Then
            lstPlayList1.Font.Bold = sFontBold
        End If
        
        If Not IsMissing(sFontItalic) Then
            lstPlayList1.Font.Italic = sFontItalic
        End If
        
        If Not IsMissing(sFontColor) And sFontColor <> -1 Then
            lstPlayList1.ForeColor = sFontColor
        End If
        
    ElseIf Object = Other Then
    
        If sFontName <> "" Then
            Label1.Font.Name = sFontName
            Label2.Font.Name = sFontName
            cbRepeat.Font.Name = sFontName
            cbShuffle.Font.Name = sFontName
        End If
        
        If sFontSize > 0 Then
            Label1.Font.Size = sFontSize
            Label2.Font.Size = sFontSize
            cbRepeat.Font.Size = sFontSize
            cbShuffle.Font.Size = sFontSize
        End If
        
        If Not IsMissing(sFontBold) Then
            Label1.Font.Bold = sFontBold
            Label2.Font.Bold = sFontBold
            cbRepeat.Font.Bold = sFontBold
            cbShuffle.Font.Bold = sFontBold
        End If
        
        If Not IsMissing(sFontItalic) Then
            Label1.Font.Italic = sFontItalic
            Label2.Font.Italic = sFontItalic
            cbRepeat.Font.Italic = sFontItalic
            cbShuffle.Font.Italic = sFontItalic
        End If
        
        If Not IsMissing(sFontColor) And sFontColor <> -1 Then
            Label1.ForeColor = sFontColor
            Label2.ForeColor = sFontColor
            cbRepeat.ForeColor = sFontColor
            cbShuffle.ForeColor = sFontColor
        End If
        
    End If
    
End Sub

Private Function SecondsToTime(lSeconds As Double) As String
'Purpose:  Changes seconds into mintues:seconds (ex.  140 becomes 2:20)
    Dim sTime As String
    Dim iSeconds As Integer
    Dim iMinutes As Integer
    
    iSeconds = Abs(Fix(lSeconds)) Mod 60
    iMinutes = Fix(Abs(Fix(lSeconds)) / 60)
    
    sTime = iMinutes & ":" & IIf(iSeconds < 10, "0", "") & iSeconds
    
    SecondsToTime = sTime
End Function

Private Function ParseFiles(sFilename As String) As Boolean
'Purpose:  Adds a song to the playlist with the file's info
    Dim sTitle As String
    Dim sArtist As String
    Dim sPlaylistName As String
    Dim sName As String
    Dim iPos As Integer
    
    If LCase(Right(sFilename, 3)) = "mp3" Then
        Mp3Info.Filename = sFilename
        
        sTitle = Mp3Info.Title
        sArtist = Mp3Info.Artist
        
        If sTitle = "" Then
            iPos = InStrRev(sFilename, "\")
            sTitle = Mid(sFilename, iPos + 1, Len(sFilename) - iPos - 4)
        End If
        
        sPlaylistName = sTitle
        
        If sArtist <> "" Then
            sPlaylistName = sPlaylistName & " - " & sArtist
        End If
        
        'lstPlayList.AddItem lstPlayList.ListCount + 1 & ". " & sPlaylistName
        lstPlayList1.ListItems.Add lstPlayList1.ListItems.Count + 1, , lstPlayList1.ListItems.Count + 1
        lstPlayList1.ListItems(lstPlayList1.ListItems.Count).ListSubItems.Add 1, , sTitle
        lstPlayList1.ListItems(lstPlayList1.ListItems.Count).ListSubItems.Add 2, , sArtist
        ParseFiles = True
    ElseIf LCase(Right(sFilename, 3)) = "mid" Or LCase(Right(sFilename, 3)) = "wav" Then
        sTitle = Mid(sFilename, iPos + 1, Len(sFilename) - iPos - 4)
        
        'lstPlayList.AddItem lstPlayList.ListCount + 1 & ". " & sTitle
        lstPlayList1.ListItems.Add lstPlayList1.ListItems.Count + 1, , lstPlayList1.ListItems.Count + 1
        lstPlayList1.ListItems(lstPlayList1.ListItems.Count).ListSubItems.Add 1, , sTitle
        lstPlayList1.ListItems(lstPlayList1.ListItems.Count).ListSubItems.Add 2, , "Unknown"
        
        ParseFiles = True
    Else
        ParseFiles = False
    End If
End Function

Private Sub LoadImages()
'Purpose:  Load the images into an imagelist, so they don't have to be re-loaded everytime a
'          bitmap is changed
    On Error GoTo EH
    ImageList1.ListImages.Clear
    ImageList1.ListImages.Add 1, "BKGRND1", LoadPicture(sImageDir & "bkgrnd1.bmp")
    ImageList1.ListImages.Add 2, "BKGRND2", LoadPicture(sImageDir & "bkgrnd2.bmp")
    ImageList1.ListImages.Add 3, "BUTTON", LoadPicture(sImageDir & "Button.bmp")
    ImageList1.ListImages.Add 4, "BUTTONDOWN", LoadPicture(sImageDir & "ButtonDown.bmp")
    ImageList1.ListImages.Add 5, "NEXTTRACK", LoadPicture(sImageDir & "NextTrack.bmp")
    ImageList1.ListImages.Add 6, "NEXTTRACKDOWN", LoadPicture(sImageDir & "NextTrackDown.bmp")
    ImageList1.ListImages.Add 7, "PAUSE", LoadPicture(sImageDir & "Pause.bmp")
    ImageList1.ListImages.Add 8, "PAUSEACTIVE", LoadPicture(sImageDir & "PauseActive.bmp")
    ImageList1.ListImages.Add 9, "PAUSEDOWN", LoadPicture(sImageDir & "PauseDown.bmp")
    ImageList1.ListImages.Add 10, "PLAY", LoadPicture(sImageDir & "Play.bmp")
    ImageList1.ListImages.Add 11, "PLAYACTIVE", LoadPicture(sImageDir & "PlayActive.bmp")
    ImageList1.ListImages.Add 12, "PLAYDOWN", LoadPicture(sImageDir & "PlayDown.bmp")
    ImageList1.ListImages.Add 13, "PREVTRACK", LoadPicture(sImageDir & "PrevTrack.bmp")
    ImageList1.ListImages.Add 14, "PREVTRACKDOWN", LoadPicture(sImageDir & "PrevTrackDown.bmp")
    ImageList1.ListImages.Add 15, "SMALLBUTTON", LoadPicture(sImageDir & "SmallButton.bmp")
    ImageList1.ListImages.Add 16, "STOP", LoadPicture(sImageDir & "Stop.bmp")
    ImageList1.ListImages.Add 17, "STOPACTIVE", LoadPicture(sImageDir & "StopActive.bmp")
    ImageList1.ListImages.Add 18, "STOPDOWN", LoadPicture(sImageDir & "StopDown.bmp")
    ImageList1.ListImages.Add 19, "TITLEBAR", LoadPicture(sImageDir & "TitleBar.bmp")
    ImageList1.ListImages.Add 20, "BKGRND3", LoadPicture(sImageDir & "bkgrnd3.bmp")
    HCSSlider1.ButtonBitmap = sImageDir & "smallbutton.bmp"
    Exit Sub
EH:
    MsgBox Err.Description & " in LoadImages"
End Sub

Private Sub SetPictures()
'Purpose:   Sets the starting pictures for all images and pictureboxes
    On Error GoTo EH

    Set imgPlay.Picture = ImageList1.ListImages(10).Picture
    Set imgPause.Picture = ImageList1.ListImages(7).Picture
    Set imgStop.Picture = ImageList1.ListImages(16).Picture
    Set imgPrev.Picture = ImageList1.ListImages(13).Picture
    Set imgNext.Picture = ImageList1.ListImages(5).Picture
    Set Picture2.Picture = ImageList1.ListImages(1).Picture
    Set Picture3.Picture = ImageList1.ListImages(2).Picture
    Set Picture4.Picture = ImageList1.ListImages(20).Picture
    
    Exit Sub
EH:
    MsgBox Err.Description & " in SetPictures"
End Sub

Private Sub PlaySong(ByVal Index As Integer)
'Purpose:  Load a song, then play it
    If Index >= 1 Then
        LoadSong Index
    Else
        LoadSong 1
    End If
    
    If MediaPlayer1.Filename <> "" Then
        MediaPlayer1.Play
        bStopPressed = False
    End If
End Sub

Private Sub LoadSong(ByVal Index As Integer)
'Purpose:  Loads a song into the mediaplayer and sets properties

    If Index <= lstPlayList1.ListItems.Count Then
        lstPlayList1.ListItems(iCurrentIndex).Bold = False
        lstPlayList1.ListItems(iCurrentIndex).ListSubItems(1).Bold = False
        lstPlayList1.ListItems(iCurrentIndex).ListSubItems(2).Bold = False
        iCurrentIndex = Index
        MediaPlayer1.Filename = lstFilenames.List(Index - 1)
        
        Label1.Caption = "Current Song: " & lstPlayList1.ListItems(Index).ListSubItems(1).Text & " by " & lstPlayList1.ListItems(Index).ListSubItems(2).Text
        'Label1.Caption = "Current Song: " & Right(lstPlayList.List(Index), Len(lstPlayList.List(Index)) - Len(Str(Index)) - 1)
        sFormTitle = frmMain.sPlayerName & " - [" & lstPlayList1.ListItems(Index).ListSubItems(1).Text & "-" & lstPlayList1.ListItems(Index).ListSubItems(2).Text & "]  "
        frmMain.Caption = sFormTitle
        HCSSlider1.Max = MediaPlayer1.Duration
        HCSSlider1.Min = 0
        lstPlayList1.ListItems(Index).Selected = True
        lstPlayList1.ListItems(Index).EnsureVisible
        lstPlayList1.ListItems(Index).Bold = True
        lstPlayList1.ListItems(Index).ListSubItems(1).Bold = True
        lstPlayList1.ListItems(Index).ListSubItems(2).Bold = True
        lstPlayList1.Refresh
    End If
End Sub

Private Sub LoadPlayList(sFilename As String)
'Purpose:  Load a playlist from a file
    Dim iFilenum As Integer
    Dim sTemp As String
    
    If Dir(sFilename) <> "" Then
        iFilenum = FreeFile
        
        Open sFilename For Input As #iFilenum
        While Not EOF(iFilenum)
            Line Input #iFilenum, sTemp
            If ParseFiles(sTemp) Then
                lstFilenames.AddItem sTemp
            End If
        Wend
        Close #iFilenum
        
        If lstPlayList1.ListItems.Count >= 1 Then
            LoadSong 1
        End If
    End If
    
End Sub

Private Sub SavePlaylist(sFilename As String)
'Purpose:  Saves the songs in the playlist to a file
    Dim iFilenum As Integer
    Dim iCnt As Integer
    
    iFilenum = FreeFile
    
    Open sFilename For Output As #iFilenum
    
    For iCnt = 1 To lstFilenames.ListCount
        Print #iFilenum, lstFilenames.List(iCnt - 1)
    Next iCnt
    
    Close #iFilenum
End Sub

Private Function GetRandomIndex(iNumOfSongs As Integer) As Integer
'Purpose:  Generates a random number for the next song
    GetRandomIndex = Int((iNumOfSongs - 0 + 1) * Rnd + 0)
End Function

Private Sub ParseMultipleFiles(sAllFiles As String)
    Dim sDir As String
    Dim iPos As String
    
    'iPos = instr(1,sallfiles,
    
End Sub

Private Sub AddDirectory(sPath As String)
    Dim p As Integer
    Dim i As Integer

    Dir1.Path = sPath
    File1.Path = sPath

    For p = 0 To Dir1.ListCount - 1
        If p < Dir1.ListCount Then
            AddDirectory Dir1.List(p)
        End If
    Next p
    For i = 0 To File1.ListCount - 1
        If LCase(Right(File1.List(i), 3)) = "mp3" _
        Or LCase(Right(File1.List(i), 3)) = "mid" _
        Or LCase(Right(File1.List(i), 3)) = "wav" Then
            If ParseFiles(Dir1.Path & "\" & File1.List(i)) Then
                lstFilenames.AddItem Dir1.Path & "\" & File1.List(i)
            End If
        End If
    Next i

    Dir1.Path = UpOneDir(Dir1.Path)
    File1.Path = Dir1.Path
End Sub

Public Function UpOneDir(sPathName As String) As String
    Dim q As Integer
    Dim num As Integer
        
    For q = 1 To Len(sPathName)
        If Mid(sPathName, q, 1) = "\" Then
            num = q
        End If
    Next q
    If Len(Mid(sPathName, 1, num - 1)) < 3 Then
        UpOneDir = Mid(sPathName, 1, num - 1) & "\"
    Else
        UpOneDir = Mid(sPathName, 1, num - 1)
    End If
End Function

Private Sub LoadINIFile()
    Dim iFilenum As Integer
    Dim sColors() As String
    Dim sTemp As String
    Dim iPos As String

    On Error GoTo EH
    
    iFilenum = FreeFile
    
    If Dir(sImageDir & "skin.ini") <> "" Then
    
        Open sImageDir & "skin.ini" For Input As #iFilenum
        
        While Not EOF(iFilenum)
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            lstPlayList1.ForeColor = CLng(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            Label1.ForeColor = CLng(sTemp)
            Label2.ForeColor = CLng(sTemp)
            cbRepeat.ForeColor = CLng(sTemp)
            cbShuffle.ForeColor = CLng(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            lstPlayList1.Font.Name = Trim$(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.Label1.Font.Name = Trim$(sTemp)
            frmMain.Label2.Font.Name = Trim(sTemp)
            frmMain.cbRepeat.Font.Name = Trim(sTemp)
            frmMain.cbShuffle.Font.Name = Trim(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.lstPlayList1.BackColor = CLng(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.BackColor = CLng(sTemp)
            cbRepeat.BackColor = CLng(sTemp)
            cbShuffle.BackColor = CLng(sTemp)
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.bConfirmDelete = CBool(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.bSavePlaylist = CBool(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.bSaveSettings = CBool(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.bScrollTitle = CBool(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.cbRepeat.Value = Val(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.cbShuffle.Value = Val(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.Slider1.Value = CInt(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.Slider2.Value = CInt(Trim(sTemp))
            
            Line Input #iFilenum, sTemp
            iPos = InStr(1, sTemp, "/")
            sTemp = Mid(sTemp, 1, iPos - 1)
            frmMain.sPlayerName = Trim(sTemp)
        Wend
        
        Close #iFilenum
    End If
    Exit Sub
EH:
    MsgBox Err.Description & " in LoadINIFile"
    Close #iFilenum
End Sub

Private Sub SaveINIFile()
    Dim iFilenum As Integer
    Dim sTemp As String
    
    iFilenum = FreeFile()
    
    If Dir(sImageDir & "skin.ini") <> "" Then
        Kill sImageDir & "skin.ini"
    End If
    
    Open sImageDir & "skin.ini" For Output As #iFilenum
    
        Print #iFilenum, lstPlayList1.ForeColor & " //Playlist Font Color"
        Print #iFilenum, frmMain.cbRepeat.ForeColor & " //All Other Font Color"
        Print #iFilenum, lstPlayList1.Font.Name & " //Playlist Font Name"
        Print #iFilenum, frmMain.cbRepeat.Font.Name & " //Other Font Name"
        Print #iFilenum, lstPlayList1.BackColor & " //Playlist Background Color"
        Print #iFilenum, frmMain.BackColor & " //All Other Background Color"
        Print #iFilenum, frmMain.bConfirmDelete & " //Confirm Delete"
        Print #iFilenum, frmMain.bSavePlaylist & " //Save Playlist"
        Print #iFilenum, frmMain.bSaveSettings & " //Save Settings"
        Print #iFilenum, frmMain.bScrollTitle & " //Scroll Title"
        Print #iFilenum, frmMain.cbRepeat.Value & " //Repeat"
        Print #iFilenum, frmMain.cbShuffle.Value & " //Shuffle"
        Print #iFilenum, frmMain.Slider1.Value & " //Volume"
        Print #iFilenum, frmMain.Slider2.Value & " //Balance"
        Print #iFilenum, frmMain.sPlayerName & " //Player Title"
    Close #iFilenum
    
    
End Sub

Function RegistryQuery(sValue As String, Optional vPrompt As Variant) As String
'Purpose: sets a value in the registry by means of input box from user
'Parameters: sValue - value to get in registry, string
'            vPrompt - text of input box, varaint, optional
    On Error GoTo ErrorHandler
    Dim sTemp As String
1    sTemp = Registry.QueryValue(HKEY_LOCAL_MACHINE, gKEYNAME, sValue, gMISSINGREGDATA)
2    If sTemp = gMISSINGREGDATA Then
3        If Not IsMissing(vPrompt) Then
4            sTemp = InputBox(vPrompt, "DealerKid")
5            Registry.SetKeyValue HKEY_LOCAL_MACHINE, gKEYNAME, sValue, sTemp, REG_SZ
6        Else
7            sTemp = gEMPTYSTRING
        End If
    End If
8    RegistryQuery = sTemp
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, "RegistryQuery " & Err.Description, Err.HelpFile, Err.HelpContext
End Function


