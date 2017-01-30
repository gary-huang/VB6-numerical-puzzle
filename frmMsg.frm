VERSION 5.00
Begin VB.Form frmMsg 
   Caption         =   "Message"
   ClientHeight    =   2085
   ClientLeft      =   3795
   ClientTop       =   5685
   ClientWidth     =   6585
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdBack 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Image imgIcon 
      Height          =   975
      Left            =   480
      Picture         =   "frmMsg.frx":0442
      Stretch         =   -1  'True
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmMsg"
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

