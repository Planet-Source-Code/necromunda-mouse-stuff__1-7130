VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   5835
   ClientTop       =   4125
   ClientWidth     =   3045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3045
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Sound"
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BOUNCE"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   960
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LOCK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SET"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RND"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lock"
      Height          =   1215
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Set"
      Height          =   1215
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1080
      Top             =   0
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mousepos.frx":0000
            Key             =   "Normal"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mousepos.frx":031A
            Key             =   "Both"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mousepos.frx":0634
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mousepos.frx":094E
            Key             =   "Right"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2280
      Picture         =   "mousepos.frx":0C68
      Top             =   1680
      Width           =   480
   End
   Begin VB.Menu menu 
      Caption         =   ""
      Begin VB.Menu mnuMini 
         Caption         =   "Minimize"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, _
     ByVal Y As Long) As Long


Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim X As Long
Dim Y As Long
Dim Xup As Boolean
Dim Yup As Boolean
Dim SW As Integer
Dim SH As Integer

Dim number As Integer
Dim P As POINTAPI
Dim LockX As Integer
Dim LockY As Integer


Private Sub Command1_Click()

    If Timer2.Enabled = False Then Timer2.Enabled = True Else Timer2.Enabled = False

End Sub

Private Sub Command2_Click()

If Text1.Text = "" And Text2.Text <> "" Then
    
    Text1.Text = 0
    SetCursorPos 0, Text2.Text

ElseIf Text2.Text = "" And Text1.Text <> "" Then
    
    Text2.Text = 0
    SetCursorPos Text1.Text, 0

Else
    SetCursorPos Text1.Text, Text2.Text

End If


End Sub

Private Sub Command3_Click()
On Error GoTo lineStop


LockX = Text3.Text
LockY = Text4.Text

If Timer3.Enabled = True Then
    
    Timer3.Enabled = False

Else

    Timer3.Enabled = True

End If

lineStop:
End Sub

Private Sub Command4_Click()

If Timer4.Enabled = True Then
    
    Timer4.Enabled = False
    
Else

    Timer4.Enabled = True


    X = 0
    Y = 0

    Xup = True
    Yup = True

    SW = Screen.Width / Screen.TwipsPerPixelX 'screen width in pixels
    SH = Screen.Height / Screen.TwipsPerPixelY 'screen height in pixels

End If
 
End Sub

Private Sub Command5_Click()

If Command5.Caption = "HIDE" Then
    Call Mouse_Hide
    Command5.Caption = "SHOW"
ElseIf Command5.Caption = "SHOW" Then
    Call Mouse_Show
    Command5.Caption = "HIDE"
End If

End Sub

Private Sub Command5_KeyDown(KeyCode As Integer, Shift As Integer)

Call Command5_Click

End Sub

Private Sub Form_Load()

Call Randomize


    menu.Visible = False
    Text1.Text = 0
    Text2.Text = 0
    Text3.Text = 0
    Text4.Text = 0
   
    

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then

Call PopupMenu(menu)

End If


End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton + vbRightButton Then
    
    Image1.Picture = ImageList1.ListImages(2).Picture
    
ElseIf Button = vbLeftButton Then

    Image1.Picture = ImageList1.ListImages(3).Picture
    
ElseIf Button = vbRightButton Then

    Image1.Picture = ImageList1.ListImages(4).Picture
    
End If


End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image1.Picture = ImageList1.ListImages(1).Picture

End Sub

Private Sub mnuMini_Click()

Me.WindowState = vbMinimized

End Sub

Private Sub Timer1_Timer()

    Cls
    Call GetCursorPos(P)

    Me.Caption = P.X & Space(1) & P.Y




End Sub

Private Sub Timer2_Timer()

Dim plus As Integer

    
    
    plus = Int(Rnd() * 2)



    Select Case plus
        Case Is = 0
         
         Call SetCursorPos(P.X + (Rnd() * 150), P.Y + (Rnd() * 150))
        
        Case Is = 1
         
         Call SetCursorPos(P.X + (Rnd() * -150), P.Y + (Rnd() * -150))
    
    End Select

End Sub

Private Sub Timer3_Timer()

 Call SetCursorPos(LockX, LockY)

End Sub

Private Sub Timer4_Timer()

If Xup = True Then
        
        X = X + 50
            
            If X >= SW Then
                
                Xup = False
        
                    If Check1.Value = 1 Then
                        beep
                    End If
        
            End If
Else
        
        X = X - 50
            
            If X <= 0 Then
                
                Xup = True
        
                   If Check1.Value = 1 Then
                        beep
                    End If
        
            End If

End If


If Yup = True Then
        
        Y = Y + 50
        
        If Y >= SH Then
            
            Yup = False
         
                 If Check1.Value = 1 Then
                        beep
                    End If
        
        End If
Else
        
        Y = Y - 50
        
        If Y <= 0 Then
        
            Yup = True
        
                If Check1.Value = 1 Then
                        beep
                    End If
        
        End If
End If
    
    Call SetCursorPos(X, Y) 'move the cursor

End Sub
