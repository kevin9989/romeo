VERSION 5.00
Begin VB.Form formulario 
   Caption         =   "Formulario"
   ClientHeight    =   11625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   11625
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Escritorio\Romeo\Todo en uno\Clientes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   9000
      Width           =   4335
   End
   Begin VB.TextBox Text5 
      DataField       =   "Observaciones"
      DataSource      =   "Data1"
      Height          =   2295
      Left            =   720
      TabIndex        =   13
      Top             =   5880
      Width           =   8055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Consulta"
      Height          =   855
      Left            =   6480
      TabIndex        =   12
      Top             =   9480
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Guardar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6240
      TabIndex        =   11
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Modificar"
      Height          =   735
      Left            =   6240
      TabIndex        =   10
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Baja"
      Height          =   735
      Left            =   6240
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alta"
      Height          =   615
      Left            =   6240
      TabIndex        =   8
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "Mail"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   7
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      DataField       =   "DNI"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   735
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      DataField       =   "Codigo"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   4
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "Observaciones"
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "DNI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "formulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Data1.Recordset.AddNew
Command1.Enabled = False
Command4.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Data1.Enabled = False


End Sub

Private Sub Command2_Click()
Dim a As String
If Data1.Recordset.EOF = False Then
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False

a = MsgBox("Esta seguro de borrar", vbYesNo + vbCritical, "Seleccionar")
If a = vbYes Then
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

Exit Sub
End If
End If




End Sub

Private Sub Command3_Click()
Data1.Recordset.Edit
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
Data1.Enabled = False

End Sub

Private Sub Command4_Click()
Data1.Recordset.Update
Command4.Enabled = False
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Data1.Enabled = True




End Sub

Private Sub Command5_Click()
formulario.Hide
base.Show
End Sub


