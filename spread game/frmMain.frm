VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Spreader [Virus Clone]"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Object of Game"
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Tag             =   $"frmMain.frx":0000
      ToolTipText     =   "Help"
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "New Game"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      ToolTipText     =   "Click to begin"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   15
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   14
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   13
      Left            =   2880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   12
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   11
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   10
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   9
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   8
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   6360
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   7
      Left            =   3600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   6
      Left            =   3240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   5
      Left            =   2880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   4
      Left            =   2520
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   3
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   2
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   1
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pbxCol 
      Height          =   255
      Index           =   0
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   6000
      Width           =   255
   End
   Begin VB.PictureBox pbxSpread 
      BackColor       =   &H00000000&
      Height          =   5805
      Left            =   120
      ScaleHeight     =   383
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   0
      Top             =   120
      Width           =   5805
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
'=Created By: Erik Stites                                                                                    =
'=Completed: Feb. 26, 2002                                                                                 =
'=Purpose: To create a fairly simple game based on the online java game      =
'=  called Virus 2.  I make no claim to the originality, in fact, Virus 2 can      =
'=  be found at: http://www.allgamesfree.com/games/                                     =
'=  This game does however show how to create and use Types, arrays,          =
'=  subroutines, nested loops, basic file I/O, testing using boolean logic,         =
'=  mesagebox, ToolTipText, and random numbers.  cmdHelp also uses the    =
'=  tag property when clicked.                                                                             =
'=  I hope that this may be useful for beginners or anyone who wants to         =
'=  make games in VB.                                                                                        =
'===============================================================

Option Explicit
'Sorry, I don't like to use the standard prefixes in my variables, I figure, knowing
' what a variable is used for is more important than the prefix; hence my naming
' the variables the way I do. Although I do use it on controls... If you notice, I use no
' timers. This is because I like event driven programming. This means that instead
' of repeating until a value is enter or changed, the code will only execute when
' something happens.

'Think of this as one box in the game
Private Type Cell
    Top As Integer
    Left As Integer
    Connected As Boolean
    Color As Long
End Type

'Used for number of cells, in this case 64 cells across, by 64 up and down
Dim CellCollection(0 To 63, 0 To 63) As Cell
Dim CellWidth As Integer, CellHeight As Integer
Dim NumCellsX As Integer, NumCellsY As Integer
Dim NumClicks As Integer, TotalConnected As Integer
Dim HighScore As Integer

Private Sub cmdHelp_Click()
    MsgBox cmdHelp.Tag, vbInformation, "A little help"
End Sub

Private Sub cmdReset_Click()
    'Click this to Begin/Start over
    Reset
    Display
End Sub

Private Sub Form_Load()
    Dim N As Integer, Xl As Integer, Yl As Integer
    Dim FName As String, FNum As Integer
    
    'Initialize Game
    'This is how many boxes Across and Down there are.
    NumCellsX = 16 'Change for more cells
    NumCellsY = 16 '   up to 64[based on variable in declarations]
    'Remember, this number is how many there are, not the highest number in the array
    
    'Self explanatory... I hope
    CellWidth = pbxSpread.ScaleWidth / NumCellsX
    CellHeight = pbxSpread.ScaleHeight / NumCellsY
    
    'Builds the color table
    For N = 0 To 15
        pbxCol(N).BackColor = QBColor(N)
        'ToolTipText is not necessary, is used for help
        pbxCol(N).ToolTipText = "Click to change to this color"
    Next N
    
    'Sets upper left corner of each cell
    For Xl = 0 To NumCellsX - 1
        For Yl = 0 To NumCellsY - 1
            CellCollection(Xl, Yl).Top = Yl * CellHeight
            CellCollection(Xl, Yl).Left = Xl * CellWidth
        Next Yl
    Next Xl
    
    'Get high score from file
    FNum = FreeFile 'Make sure we don't use some other programs open file
    If Right(App.Path, 1) = "\" Then
        FName = App.Path & "highscore.dat"
    Else
        FName = App.Path & "\highscore.dat"
    End If
    If Len(Dir(FName)) = 0 Then Exit Sub 'If file doesn't exist, skip this:
    Open FName For Input As FNum
        Input #FNum, HighScore
    Close #FNum
End Sub

Private Sub IsConnected(X As Integer, Y As Integer)
    On Error Resume Next
    'Notice the OR statement, this says that if it is already connected then it will stay that way.
    '   The AND statement means that both parts have to be true for the outcome to be true.
    CellCollection(X, Y).Connected = CellCollection(X, Y).Connected Or (CellCollection(X, Y).Color = CellCollection((X - 1), Y).Color) And CellCollection((X - 1), Y).Connected
    CellCollection(X, Y).Connected = CellCollection(X, Y).Connected Or ((CellCollection(X, Y).Color = CellCollection(X, (Y - 1)).Color) And CellCollection(X, (Y - 1)).Connected)
    CellCollection(X, Y).Connected = CellCollection(X, Y).Connected Or ((CellCollection(X, Y).Color = CellCollection(X + 1, Y).Color) And CellCollection(X + 1, Y).Connected)
    CellCollection(X, Y).Connected = CellCollection(X, Y).Connected Or ((CellCollection(X, Y).Color = CellCollection(X, Y + 1).Color) And CellCollection(X, Y + 1).Connected)
End Sub

Private Sub pbxCol_Click(Index As Integer)
    Dim Xc As Integer, Yc As Integer
    Dim ReturnValue As Integer
    
    'No this is not a mistake
    'I included two iterations of Yc to fix a color error
    'Don't worry, this doesn't change the speed by very much
    TotalConnected = 0 'used to tell if game is over
    For Xc = 0 To NumCellsX - 1
        For Yc = 0 To NumCellsY - 1
            If CellCollection(Xc, Yc).Connected = True Then
                CellCollection(Xc, Yc).Color = pbxCol(Index).BackColor
            End If
            IsConnected Xc, Yc 'Check for Adjacent
        Next Yc
        
        '============================================
        'If you comment out this block, then watch carefully...
        For Yc = 0 To NumCellsY - 1
            If CellCollection(Xc, Yc).Connected = True Then
                CellCollection(Xc, Yc).Color = pbxCol(Index).BackColor
                TotalConnected = TotalConnected + 1
            End If
            IsConnected Xc, Yc
        Next Yc
        '...you can see what the error was.
        'That is unless of course it is just my computer
        '============================================
        
    Next Xc
    'Keeps track of score
    NumClicks = NumClicks + 1
    Display
    
    'Here, we check if the game is over
    If TotalConnected = NumCellsX * NumCellsY Then
        ReturnValue = MsgBox("You cleared it in " & NumClicks & " moves." & vbCrLf & "Would you Like to play again?", vbYesNo, "Game Over")
        If ReturnValue = vbNo Then
            Unload Me
        End If
        If (NumClicks < HighScore) Or (HighScore = 0) Then
            MsgBox "You got the top score!", vbExclamation, "Congratulations!"
            HighScore = NumClicks
        End If
    End If
End Sub

Private Sub Display()
    Dim Xd As Integer, Yd As Integer
    
    'This cycles through the Cell data and displays the game board accordingly
    For Xd = 0 To NumCellsX - 1
        DoEvents
        For Yd = 0 To NumCellsY - 1
            'Draws each cell with its fill color
            pbxSpread.Line (CellCollection(Xd, Yd).Left, CellCollection(Xd, Yd).Top)-(CellCollection(Xd, Yd).Left + CellWidth, CellCollection(Xd, Yd).Top + CellHeight), CellCollection(Xd, Yd).Color, BF
        Next Yd
    Next Xd
End Sub

Private Sub Reset()
    Dim N As Integer, Xr As Integer, Yr As Integer
    
    'Clears all necessary information, randomizes the board
    NumClicks = 0
    For Xr = 0 To NumCellsX - 1
        For Yr = 0 To NumCellsY - 1
            Randomize Timer
            CellCollection(Xr, Yr).Color = QBColor(Rnd * 15)
            CellCollection(Xr, Yr).Connected = False
        Next Yr
    Next Xr
    CellCollection(0, 0).Connected = True
End Sub
