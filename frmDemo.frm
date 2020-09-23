VERSION 5.00
Begin VB.Form frmDemo 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo Form"
   ClientHeight    =   3825
   ClientLeft      =   5175
   ClientTop       =   2265
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean

Private Sub Form_Load()

On Error GoTo starterr                  'Traps error that occurs when user didn't select an option

Dim StartUpEffect

If frmOptions.optLR.Value Then          'If option Left to Right is selected
    
    StartUpEffect = &H1                 'Animates the window from left to right. This flag can be used with roll or slide animation.

ElseIf frmOptions.optRL.Value Then      'If option Right to Left is selected
    
    StartUpEffect = &H2                 'Animates the window from right to left. This flag can be used with roll or slide animation.

ElseIf frmOptions.optTB.Value Then      'If option Top to Bottom is selected
    
    StartUpEffect = &H4                 'Animates the window from top to bottom. This flag can be used with roll or slide animation.

ElseIf frmOptions.optBT.Value Then      'If option Bottom to Top is selected
    
    StartUpEffect = &H8                 'Animates the window from bottom to top. This flag can be used with roll or slide animation.

ElseIf frmOptions.optC.Value Then       'If option Centre is selected
    
    StartUpEffect = &H10                'Makes the window appear to expand outward.

Else                                    'If option Fade is selected
    
    StartUpEffect = &H80000             'Uses a fade effect. This flag can be used only if hwnd is a top-level window.

End If

If Trim(frmOptions.Text1.Text) <> "" Then   'Check if the text is spaces

AnimateWindow Me.hwnd, frmOptions.Text1.Text, StartUpEffect       'Animate Window
DoEvents

Else

MsgBox "Please enter a valid number!", , "Error"

End If

Exit Sub

starterr:                               'Error handling

MsgBox "Please select an effect!", , "Error"

End Sub
