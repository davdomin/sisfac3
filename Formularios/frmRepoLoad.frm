VERSION 5.00
Begin VB.Form frmRepoLoad 
   Caption         =   "Numero de Reposicion"
   ClientHeight    =   1755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3930
   LinkTopic       =   "Form54"
   ScaleHeight     =   1755
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumero 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   780
      Width           =   2055
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Numero de orden de Compra:"
      Height          =   435
      Left            =   780
      TabIndex        =   0
      Top             =   300
      Width           =   2295
   End
End
Attribute VB_Name = "frmRepoLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public numero
Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    numero = 0
    If KeyAscii = 13 Then numero = Val(txtNumero.Text)
    If KeyAscii = 27 Then numero = -1
    If numero <> 0 Then Unload Me
End Sub
