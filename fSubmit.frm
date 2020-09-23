VERSION 5.00
Begin VB.Form fSubmit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Submit new lyrics"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fSubmit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   345
      Left            =   3720
      TabIndex        =   17
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   16
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Optional information (fill the fields you want)"
      Height          =   1485
      Left            =   600
      TabIndex        =   8
      Top             =   3240
      Width           =   5295
      Begin VB.ComboBox cLanguage 
         Height          =   315
         ItemData        =   "fSubmit.frx":000C
         Left            =   1080
         List            =   "fSubmit.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox tYourEmail 
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   12
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox tYourName 
         Height          =   285
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   11
         Top             =   240
         Width           =   3975
      End
      Begin VB.CheckBox cShowEmail 
         Caption         =   "Show it at LYRDB.com (default NO)"
         Height          =   495
         Left            =   3000
         TabIndex        =   13
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Lyrics language:"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Your e-mail:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Your name:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mandatory fields"
      Height          =   2295
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   5295
      Begin VB.TextBox tLyrics 
         Height          =   1245
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox tTrackname 
         Height          =   285
         Left            =   1080
         MaxLength       =   155
         TabIndex        =   6
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox tArtist 
         Height          =   285
         Left            =   1080
         MaxLength       =   255
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "Lyrics:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Song title:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Artist:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Submissions are manually approved. Within 1-2 days, the lyrics will be available through LYRDB Web Services and LYRDB.com."
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   5520
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   $"fSubmit.frx":0075
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "fSubmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cSubmit As New clsLyrdb

Private Sub Command1_Click()
    Dim result As Boolean
    
    '-- We call the "UploadLyrics" method to submit the lyrics
    result = cSubmit.UploadLyrics("LyrdbSample/1.1", _
                                  tArtist.Text, _
                                  tTrackname.Text, _
                                  tLyrics.Text, _
                                  Choose(cLanguage.ListIndex + 1, _
                                         LYRDB_ENGLISH, _
                                         LYRDB_SPANISH, _
                                         LYRDB_FRENCH, _
                                         LYRDB_GERMAN, _
                                         LYRDB_ITALIAN, _
                                         LYRDB_PORTUGUESE, _
                                         LYRDB_RUSSIAN, _
                                         LYRDB_UNKNOWN), _
                                  tYourName.Text, _
                                  tYourEmail.Text, _
                                  IIf(cShowEmail.Value = Checked, True, False))
    
    '-- If "UploadLyrics" returned true, all were OK!
    If result Then
        MsgBox "Lyrics submitted successfully. Thank you very much for your collaboration. In 1-2 days the lyrics will be available for everyone through LYRDB.com and LYRDB Web Services.", vbInformation, "Success"
    
    Else
    
    '-- If not, report it.
        MsgBox "Lyrics submission error. Check your internet connection and check that you filled all the mandatory fields.", vbCritical, "Error"
    
    End If

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cLanguage.ListIndex = 7 'Language: other (default)
End Sub
