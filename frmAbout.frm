VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About"
   ClientHeight    =   9345
   ClientLeft      =   3255
   ClientTop       =   1410
   ClientWidth     =   6810
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   6810
   Begin VB.CommandButton cmdBack 
      Caption         =   "I Agree to the Above Conditions and I Agree to Pay Gary $200"
      Height          =   1095
      Left            =   720
      TabIndex        =   3
      Top             =   7920
      Width           =   5415
   End
   Begin VB.Image Image2 
      Height          =   2895
      Left            =   3600
      Picture         =   "frmAbout.frx":0742
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   240
      OLEDropMode     =   1  'Manual
      Picture         =   "frmAbout.frx":1D1B7
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   $"frmAbout.frx":4042C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "For recreational educational use only."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sliding Puzzle 1.0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title: Final Project - Sliding Puzzle v1.0
'Author: Gary Huang
'Date: May 20th, 2014
'Files: Final_Project_HuangG.frm, Final_Project_HuangG.vbp,
'       Final_Project_HuangG.bas, Final_Project_HuangG.frx,
'       frmAbout.frm, frmAbout.frx, frmMsg.frm, frmMsg.frx
'Purpose: The purpose of this program is scramble a sliding
'         puzzle board, and allow the user to solve the
'         numerical board puzzle. The high score feature
'         and graphical board will be implemented in
'         version 2.0.

Option Explicit

Private Sub cmdBack_Click()

    Unload Me

End Sub

