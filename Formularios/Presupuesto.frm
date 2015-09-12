VERSION 5.00
Begin VB.Form Form26 
   Caption         =   "Presupuesto"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   15240
   LinkTopic       =   "Form26"
   ScaleHeight     =   10950
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Tag             =   "w"
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   660
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   660
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   29
      Top             =   8880
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   28
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   9840
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   27
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   9840
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   26
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   9840
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   25
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   9840
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   24
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   23
      Top             =   8880
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cargar Presupuesto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   22
      Top             =   8880
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "+"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   21
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   20
      Top             =   8880
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   18
      Top             =   8880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   8880
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   7560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   15
      Top             =   7080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   7080
      Width           =   1575
   End
   Begin VB.ListBox List4 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   11760
      TabIndex        =   11
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   10080
      TabIndex        =   10
      Top             =   4080
      Width           =   1695
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   9360
      TabIndex        =   9
      Top             =   4080
      Width           =   735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   720
      TabIndex        =   8
      Top             =   4080
      Width           =   8655
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10080
      TabIndex        =   6
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   5
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   3720
      Width           =   8655
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   2400
      Width           =   12135
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Top             =   1920
      Width           =   7095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   2655
   End
   Begin VB.PictureBox Crystal 
      Height          =   480
      Left            =   3840
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   56
      Top             =   7560
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   55
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   54
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   53
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B50941&
      Height          =   255
      Left            =   7560
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   51
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus Trabajo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus Pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   48
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   -600
      X2              =   14640
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   -600
      X2              =   14640
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Restan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   47
      Top             =   7560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Abono"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   46
      Top             =   7080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   45
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   44
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub-Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   43
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   42
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   41
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   40
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   39
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   38
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   37
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   36
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DNI o RIF"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   35
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   -600
      X2              =   14640
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   -600
      X2              =   14640
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   -600
      X2              =   14640
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6840
      TabIndex        =   34
      Top             =   3360
      Width           =   2055
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SubTotal As Double
Public Tabla As String
Public campoClave As String
Dim esNueva As Boolean
Sub Bloqueo(st As Byte)
    If st = 1 Then 'Inicial
        Command1.Enabled = True 'Nuevo
        Command2.Enabled = False 'Guardar
        Command3.Enabled = True 'Cancelar
        Command6.Enabled = True 'Cargar Presupuesto
        Command7.Enabled = False 'Imprimir
    End If
    If st = 2 Then 'Edicion
        Command1.Enabled = False 'Nuevo
        Command2.Enabled = True 'Guardar
        Command3.Enabled = True 'Cancelar
        Command6.Enabled = True 'Cargar Presupuesto
        Command7.Enabled = True 'Imprimir
    End If
    If st = 3 Then 'Mostrar
        Command1.Enabled = True 'Nuevo
        Command2.Enabled = True 'Guardar
        Command6.Enabled = True 'Cargar Presupuesto
        Command7.Enabled = True 'Imprimir
    End If
        

End Sub
Sub ActivarBotones()
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command9.Enabled = True
    Command6.Enabled = True
    Command14.Enabled = True
    Command7.Enabled = True
    Command5.Enabled = True
End Sub
Sub CamposModificables()
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    List1.Enabled = True
    List2.Enabled = True
    List3.Enabled = True
    List4.Enabled = True
    Combo1.Enabled = True
    Text8.Enabled = True
    Text9.Enabled = True
    Command5.Enabled = True
    
    
End Sub
Sub CargarTablas()
    Tabla = "PresupuestoEnc"
    campoClave = "codPresupuesto"
End Sub
Sub cancelar()
    Limpiar
    bloquear False
End Sub
Sub Nuevo()
    Limpiar
    esNueva = True
    Text1.Text = Datos.generarCodigo("PresupuestoEnc", "codPresupuesto")
    Label18.Caption = "Pendiente"
    Label19.Caption = "Abierta"
    bloquear True
    Text4.SetFocus
    Totales
End Sub
Sub BorrarDetalle(CualDetalle)
Dim codPresupuesto As Long
Dim codProducto As Long
Dim Cantidad As Integer
Dim Monto As Double
Dim Total As Integer
Dim esServicio As Boolean

    If CualDetalle <> -1 Then
        Monto = Val(List4.List(CualDetalle))
        codPresupuesto = Val(Text1.Text)
        codProducto = Datos.MostrarCampo("Productos", "CodProducto", "Descripcion='" & List1.List(CualDetalle) & "'")
        esServicio = Datos.MostrarCampo("Productos", "Tipo", "Descripcion='" & Combo1.Text & "'") = "Servicio"
        
        Cantidad = Val(List2.List(CualDetalle))
        Conexion.Execute "delete from PresupuestoDet where codPresupuesto=" & codPresupuesto & " and codproducto =" & codProducto, Total
        If Total > 0 And Not esServicio Then
            Conexion.Execute "update Productos set stock = stock +" & Cantidad & " where codProducto=" & codProducto
        End If
        List1.RemoveItem CualDetalle
        List2.RemoveItem CualDetalle
        List3.RemoveItem CualDetalle
        List4.RemoveItem CualDetalle
        SubTotal = SubTotal - Monto
        Totales
        CuantoResta
    End If
End Sub
Sub Mostrar(rs As ADODB.Recordset)
Dim rd As New ADODB.Recordset
    Limpiar
    Text1.Text = rs(0)
    
    
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text4.Text = Datos.MostrarCampo("Clientes", "cedrif", "codCliente=" & rs(3))
    
    Text11.Text = rs(4): SubTotal = rs(4)
    Text12.Text = rs(5)
    Text13.Text = rs(6)
    'Detalle del Presupuesto
    rd.Open "select *from PresupuestoDet where CodPresupuesto=" & rs(0) & "", Conexion
    While Not rd.EOF
        List1.AddItem Datos.MostrarCampo("Productos", "Descripcion", "CodProducto=" & rd(1))
        List2.AddItem rd(2)
        List3.AddItem rd(3)
        List4.AddItem rd(4)
        rd.MoveNext
    Wend
    
    bloquear False
    Bloqueo 3
End Sub
Sub SeleccionarListas(n)
    List1.Selected(n) = True
    List2.Selected(n) = True
    List3.Selected(n) = True
    List4.Selected(n) = True
End Sub
Sub rebajarStock(Cantidad As Integer, codProducto As Integer)
    Dim iSql As String
    Dim esServicio
    esServicio = Datos.MostrarCampo("Productos", "Tipo", "codProducto=" & codProducto & "") = "Servicio"
    If Not esServicio Then
        iSql = "update Productos set stock=stock-" & Val(List2.List(i)) & " where codProducto=" & Datos.MostrarCampo("Productos", "CodProducto", "Descripcion='" & List1.List(i) & "'")
        Conexion.Execute iSql
    
    End If
End Sub
Sub GuardarDetalle(codPresupuesto As Integer)
    Dim rs As New ADODB.Recordset
    Dim contFab As Integer
    Dim iSql As String
    Dim codProducto As Integer
    contFab = 0
    For i = 0 To List1.ListCount - 1
        codProducto = Datos.MostrarCampo("Productos", "CodProducto", "Descripcion='" & List1.List(i) & "'")
        rs.Open "select *from PresupuestoDet where CodPresupuesto=" & codPresupuesto & " and codProducto=" & codProducto, Conexion
        If rs.EOF Then
            iSql = "insert into PresupuestoDet(CodPresupuesto,CodProducto,Cantidad,Precio,Total) values(" _
            & "" & codPresupuesto & "," _
            & "" & codProducto & "," _
            & "" & List2.List(i) & "," _
            & "" & List3.List(i) & "," _
            & "" & List4.List(i) & ")"
        Else
            iSql = "update PresupuestoDet set " _
            & "Cantidad=" & List2.List(i) & "," _
            & "Precio=" & List3.List(i) & "," _
            & "Total=" & List4.List(i) & " " _
            & "where codPresupuesto=" & codPresupuesto & " " _
            & "and  codProducto=" & codProducto & " "
            
            
        End If
        rs.Close
            
        Conexion.Execute iSql
    Next
    If contFab = 0 Then
    End If
End Sub
Sub guardarPago(abono As Double)
    If abono > 0 Then
        Conexion.Execute "insert into pagos (CodPago,CodPresupuesto,Fecha,Hora,CodUsuario,Monto) values(" _
        & "" & Datos.generarCodigo("Pagos", "CodPago") & "," _
        & "" & Text1.Text & "," _
        & "'" & Date & "'," _
        & "'" & Time & "'," _
        & "" & codUsuario & "," _
        & "" & abono & ")"
    End If
End Sub
Sub GuardarEncabezado()
Dim rs As New ADODB.Recordset
Dim AbonoAnterior As Double
Dim iSql As String
    GuardarCliente
    CuantoResta
    rs.Open "select *from PresupuestoEnc where codPresupuesto=" & Text1.Text, Conexion
    If rs.EOF Then
        iSql = "insert into PresupuestoEnc (codPresupuesto,Fecha,Hora,codCliente,subtotal,iva,total) values(" _
        & "" & Text1.Text & "," _
        & "'" & Date & "'," _
        & "'" & Time & "'," _
        & "" & Datos.MostrarCampo("Clientes", "codCliente", "cedrif='" & Text4.Text & "'") & "," _
        & "" & Val(Text11.Text) & "," _
        & "" & Val(Text12.Text) & "," _
        & "" & Val(Text13.Text) & ")"
        Conexion.Execute iSql
        GuardarDetalle Text1.Text
        guardarPago Val(Text14.Text)
    Else
        iSql = "Update PresupuestoEnc set " _
        & " codCliente=" & Datos.MostrarCampo("Clientes", "codCliente", "cedrif='" & Text4.Text & "'") & "," _
        & " SubTotal=" & Val(Text11.Text) & "," _
        & " Iva=" & Val(Text12.Text) & "," _
        & " Total=" & Val(Text13.Text) & " " _
        & " Where CodPresupuesto=" & Val(Text1.Text) & ""
        Conexion.Execute iSql
        If AbonoAnterior <> Val(Text14.Text) Then
            guardarPago Val(Text14.Text)
        End If
        GuardarDetalle Text1.Text
    End If

    
End Sub
Sub GuardarCliente()
    Dim rs As New ADODB.Recordset
    Dim iSql As String
    Dim Tabla As String
    Dim campoClave As String
    Dim codCliente As Integer
    Tabla = "Clientes"
    
    campoClave = "CodCliente"
    
    rs.Open "select *from clientes where cedrif='" & Text4.Text & "'", Conexion
    
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "codCliente," _
        & "cedRif," _
        & "razonsocial," _
        & "direccion," _
        & "telefono)" _
        & " values(" _
        & "" & Datos.generarCodigo(Tabla, campoClave) & "," _
        & "'" & Text4.Text & "'," _
        & "'" & Text5.Text & "'," _
        & "'" & Text6.Text & "'," _
        & "'" & Text7.Text & "')"
    Else
        codCliente = rs(campoClave)
        iSql = "UPDATE [" & Tabla & "] set " _
            & "cedrif='" & Text4.Text & "'," _
            & "RazonSocial='" & Text5.Text & "'," _
            & "direccion='" & Text6.Text & "'," _
            & "telefono='" & Text7.Text & "'" _
            & "Where codCliente=" & codCliente & ""
    End If
    Conexion.Execute iSql
    
        
End Sub
Function estaCerrada(codPresupuesto) As Boolean
Dim rs As New ADODB.Recordset
Dim vestaCerrada As Boolean
    vestaCerrada = False
    rs.Open "select *from PresupuestoEnc where codPresupuesto=" & Val(Text1.Text), Conexion
    If Not rs.EOF Then
'       vestaCerrada = rs("statusPago") = "Cerrada"
    End If
    estaCerrada = vestaCerrada

End Function
Sub CuantoResta()
    Dim Resta As Double
    Dim abono As Double
    abono = Val(Text14.Text)
    Resta = Val(Text13.Text) - abono
    Resta = Round(Resta, 2)
    If estaCerrada(Val(Text1.Text)) Then
        Resta = 0
    End If
    Text15.Text = Resta
    Label19.Caption = ObtenerStatusPago(Resta, Val(Text13.Text))

End Sub
Sub Totales()
    Dim Iva As Double
    Dim Total As Double
    Iva = SubTotal * PIVa
    Iva = Round(Iva, 2)
    Total = SubTotal + Iva
    Total = Round(Total, 2)

    Text11.Text = SubTotal
    Text12.Text = Iva
    Text13.Text = Total
    
    
    
End Sub
Function buscarLista(Descripcion As String) As Boolean
    Dim Encontro As Boolean
    Dim Pos As Integer
    Encontro = False
    Pos = 0
    While Not Encontro And Pos < List1.ListCount
        If List1.List(Pos) = Descripcion Then
            Encontro = True
        Else
            Pos = Pos + 1
        End If
    Wend
    buscarLista = Encontro
    
End Function

Function Agregar() As Boolean
Dim Agrego As Boolean
Dim Cantidad As Integer
    Agrego = False
    If Text10.Text <> "" And Text8.Text <> "" Then
        If Not buscarLista(Combo1.Text) Then
                Cantidad = Val(Int(Text8.Text))
                List1.AddItem Combo1.Text
                List2.AddItem Cantidad
                List3.AddItem Val(Text9.Text)
                List4.AddItem Text10.Text
                SubTotal = SubTotal + Val(Text10.Text)
                SubTotal = Round(SubTotal, 2)
                Combo1.SetFocus
                Text8.Text = ""
                Text9.Text = ""
                Text10.Text = ""
                Agrego = True
                Totales
        Else
            MsgBox "No puede agregar el mismo Detalle en el presupuesto"
        End If
    End If
    Agregar = Agrego
End Function
Sub CalcularMonto()
    Text10.Text = Int(Val(Text8.Text)) * Val(Text9.Text)
End Sub
Sub BuscarProducto()
    Dim filtro As String
    filtro = "Descripcion='" & Combo1.Text & "'"
    Text9.Text = Datos.MostrarCampo("Productos", "Precio", filtro)
End Sub
Sub BuscarCliente()
    Dim filtro As String
    filtro = "cedrif='" & Text4.Text & "'"
    Text5.Text = Datos.MostrarCampo("Clientes", "razonsocial", filtro)
    Text6.Text = Datos.MostrarCampo("Clientes", "direccion", filtro)
    Text7.Text = Datos.MostrarCampo("Clientes", "telefono", filtro)
End Sub
Sub CargarPresupuesto(n As Long)
    Dim rs As New ADODB.Recordset
    Dim a  As VbMsgBoxResult
    rs.Open "select *from PresupuestoEnc where codPresupuesto=" & n, Conexion
    If rs.EOF Then
        a = MsgBox("El Presupuesto " & n & " no existe", vbExclamation)
    Else
        Mostrar rs
    End If

End Sub
Sub BuscarPresupuesto()
    Dim n As Long
    n = Val(InputBox("Int Numero de Presupuesto"))
    CargarPresupuesto n

End Sub

Sub Limpiar()
    Text1.Text = ""
    Text2.Text = Date
    Text3.Text = Time
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Combo1.Text = ""
    Label18.Caption = ""
    Label19.Caption = ""
    Label23.Caption = ""
    SubTotal = 0
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
End Sub
Sub bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Text5.Enabled = es
    Text6.Enabled = es
    Text7.Enabled = es
    Text8.Enabled = es
    Text9.Enabled = es
    Text10.Enabled = es
    Text11.Enabled = es
    Text12.Enabled = es
    Text13.Enabled = es
    Text14.Enabled = es
    Text15.Enabled = es
    Text16.Enabled = es
    Combo1.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
    Label7.Enabled = es
    Label8.Enabled = es
    Label9.Enabled = es
    Label10.Enabled = es
    Label11.Enabled = es
    Label13.Enabled = es
    Label14.Enabled = es
    Label15.Enabled = es
    Label16.Enabled = es
    Label17.Enabled = es
    Label18.Enabled = es
    Label19.Enabled = es
    Label20.Enabled = es
    Label21.Enabled = es
    Label22.Enabled = es
    List1.Enabled = es
    List2.Enabled = es
    List3.Enabled = es
    List4.Enabled = es
    Command5.Enabled = es
    Command8.Enabled = es
    
End Sub

Private Sub Combo1_Click()
    BuscarProducto
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    BuscarProducto
End Sub

Private Sub Command1_Click()
    If Text4.Enabled Then
        If MsgBox("Al Hacer esta accion perdera los datos de la factura, Desea continuar ", vbYesNo) = vbYes Then
            Nuevo
            Bloqueo 2
        End If
    Else
        Nuevo
        Bloqueo 2
    End If

End Sub

Private Sub Command10_Click()
Formularios.Siguiente Me
'ActivarBotones
Bloqueo 3

End Sub

Private Sub Command11_Click()
Formularios.Ultimo Me
'ActivarBotones
Bloqueo 3

End Sub

Private Sub Command12_Click()
Formularios.Anterior Me
'ActivarBotones
Bloqueo 3

End Sub

Private Sub Command13_Click()

Formularios.Primero Me
'ActivarBotones
Bloqueo 3

End Sub

Private Sub Command14_Click()
    
    If Text4.Enabled Then
        MsgBox "Debe Tener una factura cargada"
    Else
        If Label19.Caption = "Cerrada" Then
            MsgBox "No puede modificar una factura cerrada"
        Else
            CamposModificables
            Bloqueo 2
            esNueva = False
        End If
    End If
End Sub

Private Sub Command2_Click()
    GuardarEncabezado
    bloquear False
    Bloqueo 3

End Sub
Private Sub Command3_Click()
    If Text4.Enabled Then
        If MsgBox("Al Hacer esta accion perdera los datos del presupuesto, Desea continuar ", vbYesNo) = vbYes Then
            cancelar
            Bloqueo 1
        End If
    Else
        cancelar
        Bloqueo 1

    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
On Error Resume Next
    If Not Agregar Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Command6_Click()
    BuscarPresupuesto
End Sub

Function ObtenerStatusPago(Resta As Double, Total As Double) As String
Dim Status As String
    If Resta = 0 Then
        Status = "Cerrada"
    Else
        If Resta = Total Then
            Status = "Abierta"
        Else
            Status = "Debe"
        End If
    End If
    ObtenerStatusPago = Status
End Function

Private Sub Command7_Click()
    GuardarEncabezado
    bloquear False
    Dim filtro As String
    Dim Formulas() As String
    Dim Archivo As String
    Dim codPresupuesto As Long
    codPresupuesto = Val(Text1.Text)
    filtro = "{PresupuestoEnc.CodPresupuesto}=" & codPresupuesto
    Archivo = App.Path & "\reportes\presupuesto.rpt"
    Datos.CargarReporte filtro, Archivo, Formulas
    Bloqueo 3
End Sub
Private Sub Command8_Click()
    BorrarDetalle List1.ListIndex
End Sub
Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    cancelar
    Datos.llenarCombo "select descripcion from productos order by descripcion", Combo1
    CargarTablas
    Bloqueo 1
End Sub
Private Sub List1_Click()
    SeleccionarListas List1.ListIndex
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        BorrarDetalle List1.ListIndex
    End If
End Sub
Private Sub List2_Click()
    SeleccionarListas List2.ListIndex
End Sub

Private Sub List3_Click()
    SeleccionarListas List3.ListIndex
End Sub

Private Sub List4_Click()
    SeleccionarListas List4.ListIndex
End Sub
Private Sub Text14_Change()
    CuantoResta
End Sub
Private Sub Text4_Change()
    BuscarCliente
End Sub

Private Sub Text8_Change()
    CalcularMonto

End Sub

Private Sub Text8_GotFocus()
    Text8.Text = 1
    Text8.SelStart = 0
    Text8.SelLength = 1
    
End Sub

Private Sub Text9_Change()
Dim Monto As Double
    Monto = (Val(Text9.Text) * PIVa) + Val(Text9.Text)
    Label23.Caption = "Precio+IVA:" & Round(Monto, 2)
    CalcularMonto
End Sub



