VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim reportFileName As String
Public SelectionFormula As String
Public Report As CRAXDDRT.Report
Public accion As String
Private crApp As CRAXDRT.Application
Private crxDatabaseTables As CRAXDRT.DatabaseTables
Private crxDatabaseTable As CRAXDRT.DatabaseTable
Private crxDatabase As CRAXDRT.Database
Public Sub setFileName(ByVal fileReport As String)
    reportFileName = fileReport
End Sub
Public Sub LoadExternal(externo As String)
    Set crApp = New CRAXDDRT.Application
    Set Report = crApp.OpenReport(externo)
'    RefreshData
End Sub

Private Sub RefreshData()

  Set crxDatabase = Report.Database
  Set crxDatabaseTables = crxDatabase.Tables
    For Each crxDatabaseTable In crxDatabaseTables
        crxDatabaseTable.Location = Datos.getRutaDb
    Next crxDatabaseTable
End Sub

Public Sub setParam(ByVal param As String, ByVal value As String)
    Report.ParameterFields.GetItemByName(param).AddCurrentValue (value)
End Sub
Public Sub PrintReport()
    If Report Is Nothing Then Exit Sub
    If (Report.RecordSelectionFormula = "") Then Report.RecordSelectionFormula = SelectionFormula
    If LCase(accion) = "printer" Then
        Report.PrintOut False
        Exit Sub
    End If
    FrmReporte.CRViewer91.ReportSource = Report
    FrmReporte.CRViewer91.ViewReport
    FrmReporte.CRViewer91.Zoom 90 'porcentaje
    FrmReporte.Show vbModal
End Sub
