VERSION 5.00
Begin VB.Form FormCOMPRAS 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMPRAS"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10365
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      DataField       =   "TIEMPO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      DataField       =   "CUIL/CUIT"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ALTA"
      Height          =   375
      Left            =   8160
      TabIndex        =   17
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BAJA"
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000007&
      Caption         =   "MODIFICAR"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GUARDAR"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "CANTIDAD"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "PRECIO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "PRODUCTO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "CORREO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      DataField       =   "TELEFONO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "VER COMPRAS"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   5160
      Width           =   7095
   End
   Begin VB.Data Data2 
      Caption         =   "DESPLÁCESE CON LAS FLECHAS PARA BORRAR O MODIFICAR"
      Connect         =   "Access"
      DatabaseName    =   "G:\7\Romeo\TP ROMEO\compras.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "compras"
      Top             =   4320
      Width           =   6975
   End
   Begin VB.Label Label8 
      Caption         =   "CUIL/CUIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   19
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "PRODUCTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "PRECIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "CORREO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "TELEFONO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "INGRESE SU COMPRA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "FormCOMPRAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data2.Recordset.AddNew
Data2.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False

End Sub

Private Sub Command2_Click()
If Data2.Recordset.BOF = False Then
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
A = MsgBox("¿Está seguro de borrar?", vbYesNo + vbCritical, "Seleccionar")
If A = vbYes Then
Data2.Recordset.Delete
Data2.Refresh
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Else
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End If
End If
End Sub

Private Sub Command3_Click()
If Data2.Recordset.BOF = False Then
Data2.Recordset.Edit
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
End If
End Sub

Private Sub Command4_Click()
Data2.Recordset.Update
Command4.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Data2.Enabled = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True

End Sub

Private Sub Command5_Click()
FormCOMPRAS2.Show

End Sub

