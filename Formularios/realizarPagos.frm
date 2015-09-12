VERSION 5.00
Begin VB.Form Form38 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form38"
   ClientHeight    =   9540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13395
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   45
      Charset         =   255
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form38"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   900
      Left            =   6240
      TabIndex        =   7
      Top             =   8520
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   900
      Left            =   2520
      TabIndex        =   6
      Top             =   8520
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00400000&
      Height          =   1260
      Left            =   2880
      TabIndex        =   3
      Top             =   4800
      Width           =   9855
   End
   Begin VB.Label Son 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total a pagar:"
      ForeColor       =   &H0000FFFF&
      Height          =   900
      Left            =   480
      TabIndex        =   8
      Top             =   240
      Width           =   12675
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   48
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   3000
      TabIndex        =   5
      Top             =   7320
      Width           =   9975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Devolucion:"
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   600
      TabIndex        =   4
      Top             =   6480
      Width           =   4305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pagado:"
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   600
      TabIndex        =   2
      Top             =   3600
      Width           =   5085
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   48
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   2880
      TabIndex        =   1
      Top             =   2280
      Width           =   9975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total a pagar:"
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   5475
   End
End
Attribute VB_Name = "Form38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Monto As Double
Private Sub Command1_Click()
    Monto = Val(Label2.Caption)
    Unload Me
End Sub

Private Sub Command2_Click()
    Monto = -1
    Unload Me
End Sub

Private Sub Text1_Change()
    Dim Vuelto As Double
    Vuelto = Val(Text1.Text) - Val(Label2.Caption)
    Label5.Caption = Vuelto
    
    

    
End Sub
