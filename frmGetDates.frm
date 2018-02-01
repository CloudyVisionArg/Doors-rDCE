VERSION 5.00
Begin VB.Form frmGetDates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fecha de modificacion"
   ClientHeight    =   1980
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   3324
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3324
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   900
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   360
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   900
   End
   Begin VB.TextBox txtTo 
      Height          =   288
      Left            =   960
      TabIndex        =   1
      Text            =   "txtTo"
      Top             =   840
      Width           =   1932
   End
   Begin VB.TextBox txtFrom 
      Height          =   288
      Left            =   960
      TabIndex        =   0
      Text            =   "txtFrom"
      Top             =   360
      Width           =   1932
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   192
      Left            =   360
      TabIndex        =   5
      Top             =   888
      Width           =   432
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   192
      Left            =   360
      TabIndex        =   4
      Top             =   408
      Width           =   492
   End
End
Attribute VB_Name = "frmGetDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancel As Boolean

Private Sub cmdCancel_Click()
    Cancel = True
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    ValidateDateTextbox txtFrom
    ValidateDateTextbox txtTo
    Cancel = False
    Me.Hide
End Sub

Private Sub Form_Load()
    txtFrom.Text = ""
    txtTo.Text = ""
End Sub

Private Sub txtFrom_Validate(Cancel As Boolean)
    ValidateDateTextbox txtFrom
End Sub

Private Sub txtTo_Validate(Cancel As Boolean)
    ValidateDateTextbox txtTo
End Sub

