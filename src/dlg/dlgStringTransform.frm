VERSION 5.00
Begin VB.Form dlgStringTransform 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
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
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox txtTransSeperator 
      Height          =   360
      Left            =   1380
      TabIndex        =   1
      Top             =   300
      Width           =   870
   End
   Begin VB.ComboBox cboTransString 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   330
      Width           =   975
   End
End
Attribute VB_Name = "dlgStringTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

