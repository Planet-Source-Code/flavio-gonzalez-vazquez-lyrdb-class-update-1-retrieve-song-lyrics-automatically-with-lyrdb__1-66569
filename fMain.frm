VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "LYRDB Lookup Example - by Flavio González Vázquez (flavio@ya.com)"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cSubmit 
      Caption         =   "Submit lyrics..."
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   3600
      Width           =   1695
   End
   Begin VB.OptionButton oFor 
      Caption         =   "Match"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   16
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Visit on-line documentation of LYRDB web services"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   15
      Top             =   7305
      Width           =   4845
   End
   Begin VB.CommandButton cHello 
      Caption         =   "Show LYRDB hello"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox tLyrics 
      Height          =   3975
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   1200
      Width           =   3495
   End
   Begin VB.ListBox lstResults 
      Height          =   1860
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   5535
   End
   Begin VB.CommandButton cSearch 
      Caption         =   " Search "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.OptionButton oFor 
      Caption         =   "Both (flexible)"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.OptionButton oFor 
      Caption         =   "Artist"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.OptionButton oFor 
      Caption         =   "Trackname"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox tQuery 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      Caption         =   "<-- NEW!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   3720
      Width           =   495
   End
   Begin VB.Line Line4 
      X1              =   2160
      X2              =   3240
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      X1              =   2160
      X2              =   3240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label7 
      Caption         =   $"fMain.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   8040
      Picture         =   "fMain.frx":0090
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label lError 
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   6600
      Width           =   5535
   End
   Begin VB.Label Label6 
      Caption         =   $"fMain.frx":0987
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label5 
      Caption         =   "Song lyrics (resize window to resize lyrics)"
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Double-click an item to view the lyrics"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   4575
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5640
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label3 
      Caption         =   "For:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8160
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Search:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   $"fMain.frx":0A8A
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cLyrdb As New clsLyrdb

Private Sub cHello_Click()
    If cLyrdb.ServiceAvailable Then
        MsgBox cLyrdb.LyrdbHello, vbInformation, "LYRDB Hello Message"
    Else
        MsgBox "Service unavailable! Check your internet connection", vbCritical, "Error"
    End If
End Sub

Private Sub Command1_Click()
    NavigateLyrdb
End Sub

Private Sub cSearch_Click()
    Dim query As String, field As String, n As Integer, bOk As Boolean
    query = tQuery.Text
    
    '-- First we store in a variable the type of search.
    '-- See modLyrdb for the value of the constants.
    
    If oFor(0).Value = True Then ' if "trackname" selected
        field = LYRDB_TRACKNAME
    ElseIf oFor(1).Value = True Then ' if "artist" selected
        field = LYRDB_ARTIST
    ElseIf oFor(2).Value = True Then ' if "both (flexible)" selected
        field = LYRDB_MATCH_FLEX
    ElseIf oFor(3).Value = True Then ' if "match" selected
        field = LYRDB_MATCH
    End If
    
    
    With cLyrdb
        
        '-- Before trying to access LYRDB, check if service is available
        '-- (if not, most likely that your internet connection is down)
        
        If .ServiceAvailable Then
            
            '-- We set the "agent" string to send along with the lookup.
            '-- Although it is not mandatory, is important to be set.
            '-- (no particual sintax, the string you want)
            .Agent = "LyrdbSample/1.0"
            
            '-- Make the lookup
            bOk = .LyrdbLookup(field, query)
            
            If Not bOk Then
                lError.Caption = .LastError
            Else
                lError.Caption = vbNullString
            
            
                lstResults.Clear
                
                '-- Results are stored in memory.
                '-- Now we iterate over them.
                
                For n = 0 To .Count - 1
                    '-- With the "result" index, we can access the following
                    '-- properties:
                    '    - LyrdbArtist(index)    - Returns artist for "index"
                    '    - LyrdbTrackname(index) - Returns trackname for "index"
                    '    - LyrdbID(index)        - Returns the LYRDB id for "index"
                    '                              (used to retrieve the lyrics)
                    '    - LyrdbItem(id)         - Returns the lyrics for "id"
                    '
                    '-- NOTE: Collections for Artist, Trackname and ID are
                    '         downloaded when calling to LyrdbLookup.
                    '         Lyrics are downloaded when accessing LyrdbItem
                    '         property.
                    
                    lstResults.AddItem .LyrdbID(n) & vbTab & _
                                        .LyrdbTrackname(n) & vbTab & _
                                        .LyrdbArtist(n)
                Next
                
            End If
            
        Else
            '-- Oops! Most likely that your internet connection is down.
            MsgBox "Service unavailable error. Your connection may be down...", vbCritical, "Error"
        End If
    End With
    
End Sub

Private Sub cSubmit_Click()
    fSubmit.Show 1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tLyrics.Width = ScaleWidth - tLyrics.Left ' - 120
    tLyrics.Height = ScaleHeight - tLyrics.Top ' - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Would you like to visit LYRDB documentation about the LYRDB web services?", vbYesNo) = vbYes Then
        NavigateLyrdb
    End If
End Sub

Sub NavigateLyrdb()
    Shell "explorer http://www.lyrdb.com/services/about-lws.php", vbMaximizedFocus
End Sub

Private Sub lstResults_DblClick()
    Dim lyrics As String, id As String
    
    '-- Download the lyrics passing the lyrics ID to
    '-- LyrdbItem function.
    
    id = Left(lstResults.Text, InStr(lstResults.Text, vbTab))
    lyrics = cLyrdb.LyrdbItem(id)
    tLyrics.Text = lyrics
    
End Sub
