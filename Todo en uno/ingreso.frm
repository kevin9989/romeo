VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form ingreso 
   Caption         =   "Ingreso"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "VALIDAR"
      Height          =   1215
      Left            =   2880
      TabIndex        =   5
      Top             =   5160
      Width           =   2775
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   1080
      TabIndex        =   4
      Top             =   3960
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1296
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   3240
   End
   Begin VB.TextBox Text2 
      Height          =   500
      Left            =   3000
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   500
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "100"
      Height          =   255
      Left            =   7920
      TabIndex        =   7
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "ingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim a As String
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 99 Then
Timer1.Enabled = False
If Text1.Text = "admin" And Text2.Text = "123" Then
ingreso.Hide
formulario.Show
Else
a = MsgBox("Usuario y/o Contraseña incorrecta", vbOKOnly + vbCritical, "Eror de validación")



End If
End If




End Sub
