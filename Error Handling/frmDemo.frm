VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.XPButton XPButton1 
      Height          =   375
      Left            =   870
      TabIndex        =   0
      Top             =   2520
      Width           =   2700
      _extentx        =   4763
      _extenty        =   661
      caption         =   "Show Demo Error Msg"
      font            =   "frmDemo.frx":0000
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title        : Reusable Error Msgbox
'module name          : frmDemo
'Version              : 1.0
'Programmer Name      : Amod Gokhale
'Date of Creation     : 01 Aug 2005
'Date of Modification : 01 Aug 2005
Option Explicit

Private Sub Command1_Click()
End Sub

Private Sub XPButton1_Click()
'below code is demo application to show errormsg
On Error GoTo ErrorHandler
    Dim Test As ErrObject
    
    Test.Raise 12
    Exit Sub
ErrorHandler:
    'Below 2 methods should be called in each error handling of VB
    Call frmErrorMsg.SetErrorLog(Err.Number, Err.Description, "Command1_Click", "frmDemo")
    frmErrorMsg.Show vbModal
    End

End Sub
