VERSION 5.00
Begin VB.Form Form37 
   Caption         =   "Calculadora de Costos"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form37"
   ScaleHeight     =   5760
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   6000
      TabIndex        =   15
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2160
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1500
      Left            =   3360
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   3360
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Calcular Costo"
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
      Left            =   2760
      TabIndex        =   16
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label5"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Precio de Coste Indiividual"
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label5"
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Total:"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Precio:"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Material"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Piezas del Lote:"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "Form37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Total As Double
Public Resultado As Double
Sub MOntos()
    Dim Cantidad As Double
    Cantidad = Val(Text1.Text)
    Label5.Caption = Total
    If Cantidad > 0 Then
        Label7.Caption = Total / Cantidad
    End If
End Sub

Private Sub Command1_Click()
Resultado = Val(Label7.Caption)
Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    List1.AddItem Text2.Text
    List2.AddItem Val(Text3.Text)
    Total = Total + Val(Text3.Text)
MOntos
End Sub

Private Sub Command4_Click()
Dim A As Integer
    A = List1.ListIndex
    If A <> -1 Then
        Total = Total - Val(List2.List(A))
        List1.RemoveItem A
        List2.RemoveItem A
        MOntos
    Else
        MsgBox "Debe elegir un material"
    End If
    
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    Text1.Text = ""
    Total = 0

End Sub

Private Sub Text1_Change()
    MOntos
End Sub
