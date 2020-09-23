VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   227
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtTone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "35"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tone (0-127):"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub txtTone_Validate(Cancel As Boolean)
    If txtTone.Text = "" Or _
        Val(txtTone.Text) < 0 Or _
        Val(txtTone.Text) > 127 Then
            Cancel = True
            Beep
    End If
End Sub
