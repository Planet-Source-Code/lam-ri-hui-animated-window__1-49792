VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3825
   ClientLeft      =   9945
   ClientTop       =   2265
   ClientWidth     =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   1815
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   9
      Text            =   "5000"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Effect"
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton optC 
         Caption         =   "Centre"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Tag             =   "&H4"
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton optBT 
         Caption         =   "Bottom to Top"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Tag             =   "&H4"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optTB 
         Caption         =   "Top to Bottom"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Tag             =   "&H3"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optRL 
         Caption         =   "Right to Left"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Tag             =   "&H2"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optLR 
         Caption         =   "Left to Right"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Tag             =   "&H1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optFade 
         Caption         =   "Fade"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Time (in milliseconds)"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1485
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'*****  Animated Window  *****
'*****************************

'Add special effects when you are showing a form!
'Animated Window demonstrates how to add special
'effect when showing a form.

'But make sure you have Windows 2000 and later &
'Windows 98 and later to apply these effect.

'If you like this code, don't forget to vote for me
'at www.planetsourcecode.com/vb/

'HAPPY CODING!

Option Explicit

Private Sub Command1_Click()
Unload frmDemo          'Unload the form if the form is loaded
Pause (0.2)             'Pause the program for 0.2 second
Load frmDemo            'Show the form
frmDemo.SetFocus        'Set the focus to the form
End Sub

Sub Pause(Interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(Interval)
DoEvents
Loop
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmDemo          'Unload frmDemo as well when the option form unloaded
End Sub
