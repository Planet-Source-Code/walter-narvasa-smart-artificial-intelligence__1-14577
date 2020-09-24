VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "ANIGIF.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   " Funny Jo3  /  By Soldier007"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8400
      Top             =   7200
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Voice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   11040
      Width           =   975
   End
   Begin AniGIFCtrl.AniGIF AniGIF1 
      Height          =   1215
      Left            =   3240
      TabIndex        =   6
      Top             =   6120
      Width           =   2055
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   2
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":0000
      ExtendWidth     =   3625
      ExtendHeight    =   2143
      Loop            =   0
      AutoRewind      =   0   'False
      Synchronized    =   -1  'True
   End
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   11535
      Left            =   14880
      ScaleHeight     =   11535
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   0
      Width           =   465
   End
   Begin VB.ListBox Used 
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   11040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7680
      TabIndex        =   0
      Top             =   11040
      Width           =   4815
   End
   Begin AniGIFCtrl.AniGIF AniGIF3 
      Height          =   1215
      Left            =   6480
      TabIndex        =   7
      Top             =   5880
      Width           =   2055
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   2
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   2
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":3C78
      ExtendWidth     =   3625
      ExtendHeight    =   2143
      Loop            =   0
      AutoRewind      =   0   'False
      Synchronized    =   -1  'True
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   5280
      Top             =   3840
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type Your Question :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   7680
      TabIndex        =   4
      Top             =   10680
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   11535
      Left            =   0
      Picture         =   "Form1.frx":78F0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cL As New cLogo
Dim Peedy As IAgentCtlCharacterEx
Const DATAPATH = "C:\WINDOWS\MSAGENT\CHARS\PEEDY.ACS"
Dim Request As IAgentCtlRequest


Function OldQuestion(UseText As String) As Boolean
    Dim i
    For i = 0 To Used.ListCount - 1
        If UCase(Used.List(i)) = UCase(UseText) Then
            OldQuestion = True
            Exit Function
        Else
            OldQuestion = False
        End If
    Next i
End Function


Function OkQuestion(TheText As String)
    Dim TempText As String
    Dim Ekstra As String
    Dim Text(0 To 15) As String
    Dim Number As Integer
    If InStr(1, TheText, "elvis", vbTextCompare) Then GoTo Theking
        TempText = Replace(TheText, " ", "")
    If TheText = TempText Then
        Ekstra = ""
    Else
        Ekstra = " "
    End If
Start:
    If InStr(1, TheText, "What" & Ekstra, vbTextCompare) Then GoTo WhichWhatHow
    If InStr(1, TheText, "How" & Ekstra, vbTextCompare) Then GoTo WhichWhatHow
    If InStr(1, TheText, "Where" & Ekstra, vbTextCompare) Then GoTo Where
    If InStr(1, TheText, "Why" & Ekstra, vbTextCompare) Then GoTo Why
    If InStr(1, TheText, "Which" & Ekstra, vbTextCompare) Then GoTo WhichWhatHow
    If InStr(1, TheText, "Who" & Ekstra, vbTextCompare) Then GoTo Who
    If InStr(1, TheText, "When" & Ekstra, vbTextCompare) Then GoTo When

    Text(0) = "I think so"
    Text(1) = "What the hell do you think ?"
    Text(2) = "Yeah"
    Text(3) = "Nope"
    Text(4) = "Comeone .. evryone knows the answer is NO"
    Text(5) = "Yepper"
    Text(6) = "Not as far as i know"
    Text(7) = "Yes"
    Text(8) = "No"
    Text(9) = "Comeone .. evryone knows the answer is Yes"
    Text(10) = "Yeah Right?"
    Text(11) = "Are you kiddin' me?"
    Text(12) = "Don't try to make me laugh!"
    Text(13) = "Ok then, Will you further more explain to me?"
    Text(14) = "You have delayed reaction on what you are saying!"
    Text(15) = "Don't ask me!, Ask your Dog"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
    Exit Function
WhichWhatHow:
    Text(0) = "How the hell should i know ?"
    Text(1) = "It's getting late ... ask me tomorrow"
    Text(2) = "Ehhhhhh"
    Text(3) = "Weeelllll"
    Text(4) = "Thats funny ..... someone just asked me that lately"
    Text(5) = "Damn thats's a good question"
    Text(6) = "You would like to know that i think ;)"
    Text(7) = "I won't tell you that"
    Text(8) = "Damn i forgot"
    Text(9) = "I wish you knew how much i hate curious people"
    Text(10) = "Are you sure with your question?"
    Text(11) = "Until further notice my friend!"
    Text(12) = "Are you making me feel guilty?"
    Text(13) = "Just as what everybody does"
    Text(14) = "I will answer that in the near future OK!"
    Text(15) = "That is pretty personal question, I may say"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
    Exit Function
Where:
    Text(0) = "Wasn't that in florida ?"
    Text(1) = "I think it's somehwere around America"
    Text(2) = "Just a second .. i'll go get the atlas"
    Text(3) = "In Germany"
    Text(4) = "Right up in heaven"
    Text(5) = "In your grandmoms backyard"
    Text(6) = "I think it was in the tomato-soup we got last friday"
    Text(7) = "It's in the whitehouse"
    Text(8) = "Last time i cheked it was in my underwear"
    Text(9) = "Maybe in my new washmachine ?"
    Text(10) = "In my bed"
    Text(11) = "Right down in Hell!!"
    Text(12) = "At you place"
    Text(13) = "Your Dog House!"
    Text(14) = "Right in front of you Bath Room"
    Text(15) = "At the lake"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
Exit Function
Why:
    Text(0) = "Because you ask all those stupid questions"
    Text(1) = "Check in the encyclopedia"
    Text(2) = "It was my destiny"
    Text(3) = "Ehmmmm"
    Text(4) = "because that's the way i wanted it"
    Text(5) = "Ask Bill Clinton"
    Text(6) = "Weellllll"
    Text(7) = "I don't remember"
    Text(8) = "Come on ask me a relevant question"
    Text(9) = "Why don't you ask your mom ?"
    Text(10) = "I have amnesia right now can't remember anything!!"
    Text(11) = "because you try to kiss my ass"
    Text(12) = "I don't care!! The hell with you!!"
    Text(13) = "because we are meant for each other"
    Text(14) = "because you always makes me horny??!!!"
    Text(15) = "I beg your pardon, Please come again?"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
    Exit Function
Who:
    Text(0) = "Don't you think it was Donald Duck"
    Text(1) = "James Bond"
    Text(2) = "Jim Carry"
    Text(3) = "A strange man in the middle of the 80's"
    Text(4) = "Mrs. Winterborne"
    Text(5) = "It was your dentist"
    Text(6) = "Luke Perry"
    Text(7) = "Your dad"
    Text(8) = "Bill Clinton"
    Text(9) = "Pamela Andersons big tittys"
    Text(10) = "Dennis Rodman"
    Text(11) = "Rosanna Roces"
    Text(12) = "Ara Mina"
    Text(13) = "Your Mother, Ok!"
    Text(14) = "Your Priest"
    Text(15) = "My snake"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
    Exit Function
When:
    Text(0) = "Tomorrow"
    Text(1) = "Yesterday"
    Text(2) = "It was in year 1900"
    Text(3) = "When you are old enough to hear it"
    Text(4) = "Damn it's a long time ago"
    Text(5) = "In the Vietnam war"
    Text(6) = "When the day you were born"
    Text(7) = "The first day your father changed his underwear"
    Text(8) = "I't was in the same period when Bill Clinton and Monica ........."
    Text(9) = "Now"
    Text(10) = "The day after tomorrow"
    Text(11) = "Last Saturday"
    Text(12) = "Tomorrow Evening!"
    Text(13) = "Fucking Monday!!, Ok!"
    Text(14) = "Since the first day I fuck your dog."
    Text(15) = "Last Friday, when I lick you cats pussy!"
    Number = Int((Rnd * 16) + 1) - 1
    OkQuestion = Text(Number)
    Exit Function
Theking:
    If InStr(1, TheText, "alive", vbTextCompare) Or InStr(1, TheText, "living", vbTextCompare) Or InStr(1, TheText, "dead", vbTextCompare) Then
        OkQuestion = "What the hell do you think ? Offcourse he's alive, remember he's the king ;)"
    Else
        GoTo Start
    End If
End Function

Function NoQuestion()
    Dim Text(0 To 15) As String
    Dim Number As Integer
    Text(0) = "Ehh... why just try with a question ;)"
    Text(1) = "I'm better to answer questions"
    Text(2) = "Wee OK ... well what about asking me a question"
    Text(3) = "I love questions with sense only!"
    Text(4) = "Pllzz ask me a question"
    Text(5) = "I'm only a bird .... ask me a question"
    Text(6) = "Try me with a question ;)"
    Text(7) = "OK, but ask me a question now"
    Text(8) = "Let's try a with a fucking question"
    Text(9) = "I'm only good at questions"
    Text(10) = "Are you running out of questions?"
    Text(11) = "Maybe you are sick of me thats why you don't want to ask me some questions"
    Text(12) = "I think you are dumb thats why you can't ask me a question!"
    Text(13) = "I'm sick and tired of what are saying, just ask me questions please!"
    Text(14) = "Questions only or else i will kick you out!!"
    Text(15) = "I'm sure you are not only dumb but silly also, ask me a good question!"
    Randomize
    Number = Int((Rnd * 16) + 1) - 1
    NoQuestion = Text(Number)
End Function

Function AI(Text As String)
    Dim TempText As String
    Dim Ekstra As String
    TempText = Replace(Text, " ", "")
    If OldQuestion(TempText) = False Then
        Used.AddItem TempText
        If Text = TempText Then
            Ekstra = ""
        Else
            Ekstra = " "
        End If
        If InStr(1, Text, "What" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "How" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "Where" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "Why" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "Which" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "Who" & Ekstra, vbTextCompare) Then GoTo Question
        If InStr(1, Text, "When" & Ekstra, vbTextCompare) Then GoTo Question
        If Right(TempText, 1) = "?" Then GoTo Question
        AI = NoQuestion
        Exit Function
Question:
        AI = OkQuestion(Text)
    Else
        AI = "I already answered you on that.."
    End If
End Function

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    MsgBox "Plz try"
End Sub


Private Sub Command2_Click()
    If Command2.Caption = "Voice" Then
        Command2.Caption = "Cancel"
        Timer1.Enabled = True
    Else
        Command2.Caption = "Voice"
        Timer1.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    cL.DrawingObject = picLogo
    cL.Caption = "Artificial Intelligence by Walter A. Narvasa"
    cL.StartColor = vbBlue
    cL.EndColor = vbBlack
    Agent1.Characters.Load "Peedy", DATAPATH
    Set Peedy = Agent1.Characters("Peedy")
    Peedy.LanguageID = &H409
    Peedy.Show
    Call AddPeedyCommands
    Peedy.MoveTo 34, 109
    Peedy.Speak "Welcome to Peedy's Question and Answer!"
    Peedy.Speak "I am an Artificial Intelligence being and was created by Walter A. Narvasa"
    Peedy.Speak "You can ask me any question you want even my love life."
    Peedy.Speak "But ask me in a nice manner and in English please because i can't undestand Tagalog, Ok!"
    Peedy.Speak "hek hek hek hek hek hek hek hek.."
    Peedy.MoveTo 389, 322
    Peedy.Play "Surprised"
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    picLogo.Height = Me.ScaleHeight
    On Error GoTo 0
    cL.Draw
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Peedy.Speak AI(Text1.Text)
        Text1.Text = ""
    End If
End Sub


Private Sub AddPeedyCommands()
    'Add voice commands to Peedy a.k.a. (Peedy)
    Peedy.Commands.Add "what", "what", "what", True, True
    Peedy.Commands.Add "how", "how", "how", True, True
    Peedy.Commands.Add "where", "where", "where", True, True
    Peedy.Commands.Add "why", "why", "why", True, True
    Peedy.Commands.Add "which", "which", "which", True, True
    Peedy.Commands.Add "who", "who", "who", True, True
    Peedy.Commands.Add "when", "when", "when", True, True
    Peedy.Commands.Add "doyou", "do you", "do you", True, True
    Peedy.Commands.Add "goodbye", "Good By", "Good By", True, True
    Peedy.Commands.Add "exit", "exit", "exit", True, True
End Sub

Private Sub Agent1_Click(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Integer, ByVal y As Integer)
    If Button = vbLeftButton Then
        Peedy.Play "Surprised"
        Peedy.Speak "Be careful with that pointer!|Don't touch me!|OUCH!|Don't try to fondle me!|Get back to what you are doing!"
        Peedy.Play "RestPose"
    End If
End Sub

Private Sub Agent1_Command(ByVal UserInput As Object)
    Select Case UserInput.Name
    Case "what"
            Peedy.Speak AI("what")
    Case "how"
            Peedy.Speak AI("how")
    Case "where"
            Peedy.Speak AI("where")
    Case "why"
            Peedy.Speak AI("why")
    Case "which"
            Peedy.Speak AI("which")
    Case "what"
            Peedy.Speak AI("what")
    Case "who"
            Peedy.Speak AI("who")
    Case "when"
            Peedy.Speak AI("when")
    Case "doyou"
            Peedy.Speak AI("do you")
    Case "goodbye"
        Peedy.Speak "You said Good bye to Peedy"
    Case "exit"
        Peedy.Speak "Peedy is ready to exit, Bye!"
  End Select
End Sub

Private Sub Timer1_Timer()
    Peedy.Listen (True)
End Sub
