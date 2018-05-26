VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONTROL"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FF00&
      Caption         =   "CLIENTE"
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "INVENTARIO"
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "SALIR"
      Height          =   615
      Left            =   7920
      TabIndex        =   3
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "EMPLEADOS"
      Height          =   615
      Left            =   7800
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "COMPRAS"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "VENTAS"
      Height          =   615
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

FormVENTAS.Show
End Sub

Private Sub Command2_Click()

FormCOMPRAS.Show
End Sub

Private Sub Command3_Click()

FormEMPLEADOS.Show
End Sub

Private Sub Command5_Click()

FormINVENTARIO.Show
End Sub

Private Sub Image1_Click()

End Sub
