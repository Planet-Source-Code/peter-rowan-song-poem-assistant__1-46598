VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Poem/Song Writter"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOpen 
      Caption         =   "Select WordList"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Word List"
      Filter          =   "*.Txt"
   End
   Begin VB.TextBox Txtchars 
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Text            =   "3"
      Top             =   480
      Width           =   855
   End
   Begin VB.ListBox ListResults 
      Height          =   3765
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Txtword 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Last Characters, eg last 3 characters of hello would return all words that end in 'llo'"
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Rhyming Word"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Search As String
Dim WordList As String



Private Sub CmdOpen_Click()

CD.ShowOpen 'open the box thing
WordList = CD.FileName
End Sub

Private Sub CmdStart_Click()

If WordList = "" Then
    MsgBox "No Word List selected"
    Exit Sub
    
ElseIf Txtword.Text = "" Then
    MsgBox "No Search Selected"
    Exit Sub
    
End If


ListResults.Clear 'clear the listbox

Search = Txtword.Text 'set up the search variable, this is what you type in
SearchPattern = Right$(Search, Txtchars.Text) 'work out last characters of input word



 Open WordList For Input As #1 'open the wordlist



Do While Not EOF(1) ' Loop until end of file.

    Line Input #1, NewWord ' Read word into variable, file has to have 1 word per line
       
        
    Pattern = Right$(NewWord, Txtchars.Text) 'find out last letters of new word, eg. last 3 letters of hello would be 'ello'
   
    
    
    
    
    If Pattern = SearchPattern Then ListResults.AddItem NewWord 'check if the same ending pattern, and add to list
        
  
    
Loop 'keep looping until end of file
Close #1    ' Close file.

MsgBox "Finished" 'tell user it's done.

End Sub
