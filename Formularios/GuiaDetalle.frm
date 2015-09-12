VERSION 5.00
Begin VB.Form Form51 
   Caption         =   "Datos de la guia de Ventas"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11130
   LinkTopic       =   "Form51"
   ScaleHeight     =   6825
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   495
      Left            =   9480
      TabIndex        =   17
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Datos de la guia de Ventas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Haga doble click para expandir la pantalla"
      Top             =   360
      Width           =   12735
      Begin VB.TextBox Text21 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Width           =   9495
      End
      Begin VB.TextBox Text22 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Width           =   6375
      End
      Begin VB.TextBox Text23 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox Text24 
         Height          =   315
         Left            =   5880
         TabIndex        =   3
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox Text25 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   2040
         Width           =   9495
      End
      Begin VB.TextBox Text26 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox Text27 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   3120
         Width           =   9495
      End
      Begin VB.TextBox Text28 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   3720
         Width           =   9495
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Señores"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Tranpor."
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Destino"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "RUC"
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Placa:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Partida"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Llegada"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form51"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form10.Text21.Text = Me.Text21.Text
    Form10.txtTranportista.Text = Me.Text22.Text
    Form10.Text23.Text = Me.Text23.Text
    Form10.Text24.Text = Me.Text24.Text
    Form10.Text25.Text = Me.Text25.Text
    Form10.Text26.Text = Me.Text26.Text
    Form10.Text27.Text = Me.Text27.Text
    Form10.Text28.Text = Me.Text28.Text
    Unload Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me

End Sub

