VERSION 5.00
Begin VB.Form FormEMPLEADOS 
   BackColor       =   &H8000000D&
   Caption         =   "Form4"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11910
   LinkTopic       =   "Form4"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text14 
      DataField       =   "TIEMPO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text13 
      DataField       =   "TURNO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   4680
      Width           =   3015
   End
   Begin VB.TextBox Text12 
      DataField       =   "SEXO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox Text11 
      DataField       =   "NACIONALIDAD"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   6480
      Width           =   3255
   End
   Begin VB.TextBox Text10 
      DataField       =   "CARGO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      DataField       =   "AREA"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "VER EMPLEADOS"
      Height          =   495
      Left            =   3360
      TabIndex        =   21
      Top             =   7920
      Width           =   7095
   End
   Begin VB.Data Data2 
      Caption         =   "DESPLÁCESE CON LAS FLECHAS PARA BORRAR O MODIFICAR"
      Connect         =   "Access"
      DatabaseName    =   "G:\7\Romeo\TP ROMEO\empleados2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "empleados2"
      Top             =   7200
      Width           =   6975
   End
   Begin VB.TextBox Text8 
      DataField       =   "FECHA DE NACIMIENTO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "CUIL"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   2940
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5880
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "CORREO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   2940
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      DataField       =   "NOMBRE"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      DataField       =   "DIRECCION"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   2940
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5280
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GUARDAR"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MODIFICAR"
      Height          =   375
      Left            =   9720
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BAJA"
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ALTA"
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      DataField       =   "DNI"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   2940
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   4080
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      DataField       =   "APELLIDO"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label14 
      Caption         =   "TURNO"
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
      Left            =   6600
      TabIndex        =   31
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "SEXO"
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
      Left            =   1080
      TabIndex        =   30
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "NACIONALIDAD"
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
      TabIndex        =   27
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "CARGO"
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
      Left            =   6600
      TabIndex        =   26
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "AREA"
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
      Left            =   6600
      TabIndex        =   25
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "FECHA DE NACIMIENTO"
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
      TabIndex        =   20
      Top             =   3480
      Width           =   2775
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
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "CUIL"
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
      Left            =   960
      TabIndex        =   17
      Top             =   5880
      Width           =   1215
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
      Left            =   960
      TabIndex        =   16
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "NOMBRE"
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
      Left            =   1200
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "DIRECCION"
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
      Left            =   960
      TabIndex        =   14
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "INGRESE LOS DATOS DEL NUEVO EMPLEADO"
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
      Left            =   2400
      TabIndex        =   13
      Top             =   120
      Width           =   7935
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
      Left            =   960
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "APELLIDO"
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
      Left            =   1080
      TabIndex        =   11
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "FormEMPLEADOS"
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
Text8.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Text13.Locked = False


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
Text8.Locked = False
Text9.Locked = False
Text10.Locked = False
Text11.Locked = False
Text12.Locked = False
Text13.Locked = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = True
End If
End Sub
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
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
Text13.Locked = True
End Sub

Private Sub Command6_Click()
FormEMPLEADOS2.Show
End Sub


