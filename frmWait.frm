VERSION 5.00
Begin VB.Form frmWait 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Shape Shape1 
      Height          =   625
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning File"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   340
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait.... "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2055
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
