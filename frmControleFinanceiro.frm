VERSION 5.00
Begin VB.Form frmControleFinanceiro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle Financeiro"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12900
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmControleFinanceiro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   12900
End
Attribute VB_Name = "frmControleFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Private Item                  As ListItem

'Private Sub MontaLvw()
'   lvwConsulta.ListItems.Clear
'   lvwConsulta.ColumnHeaders.Clear
'   lvwConsulta.Gridlines = True
'   lvwConsulta.ColumnHeaders.Add , , "", 300
'   lvwConsulta.ColumnHeaders.Add , , "ID", 900
'   lvwConsulta.ColumnHeaders.Add , , "Tipo", 4000
'   lvwConsulta.ColumnHeaders.Add , , "Descrição", 7000
'   lvwConsulta.View = lvwReport
'End Sub
'
'Private Sub Form_Load()
'   MontaLvw
'End Sub



