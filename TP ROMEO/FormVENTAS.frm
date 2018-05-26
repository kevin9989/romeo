VERSION 5.00
Begin VB.Form FormVENTAS 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VENTAS"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10905
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      DataField       =   "TIEMPO"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      DataField       =   "CANTIDAD"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      DataField       =   "DNI"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2640
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ALTA"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BAJA"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MODIFICAR"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GUARDAR"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "DESPLÁCESE CON LAS FLECHAS PARA BORRAR O MODIFICAR"
      Connect         =   "Access"
      DatabaseName    =   "G:\7\Romeo\TP ROMEO\ventas.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ventas"
      Top             =   4560
      Width           =   6975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "VER VENTAS"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   5160
      Width           =   7095
   End
   Begin VB.TextBox Text6 
      DataField       =   "PRECIO"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      DataField       =   "PRODUCTO"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "CORREO"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "CLIENTE"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label8 
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
      Left            =   360
      TabIndex        =   19
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "D.N.I"
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
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "INGRESE SU VENTA"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label6 
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
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Left            =   5520
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "CLIENTE"
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
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
End
Attribute VB_Name = "FormVENTAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim CID As Integer
Dim A As String
CID = CID + 1
Text1.Text = CID
End Sub

Private Sub Command1_Click()
Data1.Recordset.AddNew
Data1.Enabled = False
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
If Data1.Recordset.BOF = False Then
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
A = MsgBox("¿Está seguro de borrar?", vbYesNo + vbCritical, "Seleccionar")
If A = vbYes Then
Data1.Recordset.Delete
Data1.Refresh
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
If Data1.Recordset.BOF = False Then
Data1.Recordset.Edit
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
Data1.Recordset.Update
Command4.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Data1.Enabled = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
End Sub

Private Sub Command5_Click()
FormVENTAS2.Show

End Sub

