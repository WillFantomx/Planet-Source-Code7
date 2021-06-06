VERSION 5.00
Begin VB.Form frmErrorMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Your Application Name"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00D6E7EF&
      Height          =   2625
      Left            =   -45
      ScaleHeight     =   2565
      ScaleWidth      =   6255
      TabIndex        =   3
      Top             =   990
      Width           =   6315
      Begin Project1.XPButton XPButton2 
         Height          =   360
         Left            =   5130
         TabIndex        =   9
         Top             =   2115
         Width           =   990
         _extentx        =   1746
         _extenty        =   635
         caption         =   "&Don't send"
         font            =   "Form1.frx":0000
      End
      Begin Project1.XPButton XPButton1 
         Height          =   360
         Left            =   3180
         TabIndex        =   8
         Top             =   2115
         Width           =   1920
         _extentx        =   3387
         _extenty        =   635
         caption         =   "&Send Error Report"
         font            =   "Form1.frx":0028
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D6E7EF&
         Caption         =   "Please tell ""Your application name"" about this problem."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   690
         Width           =   5475
      End
      Begin VB.Label lblDescript 
         BackColor       =   &H00D6E7EF&
         Caption         =   "type your text in form load method"
         Height          =   600
         Left            =   180
         TabIndex        =   6
         Top             =   990
         Width           =   5595
      End
      Begin VB.Label Label4 
         BackColor       =   &H00D6E7EF&
         Caption         =   "To see what data this error report contains,"
         Height          =   240
         Left            =   180
         TabIndex        =   5
         Top             =   1740
         Width           =   3090
      End
      Begin VB.Label Label5 
         BackColor       =   &H00D6E7EF&
         Caption         =   "click here"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3330
         MouseIcon       =   "Form1.frx":0050
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1740
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00D6E7EF&
         Caption         =   "A log of this error has been created."
         Height          =   405
         Left            =   180
         TabIndex        =   4
         Top             =   210
         Width           =   5595
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   -45
      ScaleHeight     =   945
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6435
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "The application has recovered from serious problem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   2
         Top             =   345
         Width           =   5550
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6165
      Y1              =   990
      Y2              =   990
   End
End
Attribute VB_Name = "frmErrorMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project Title        : Reusable Error Msgbox
'module name          : frmErrorMsgbox
'Version              : 1.0
'Programmer Name      : Amod Gokhale
'Date of Creation     : 01 Aug 2005
'Date of Modification : 01 Aug 2005
Option Explicit

Dim strErrDesc As String
Dim nErrNo As Long
Dim strMethodName As String
Dim strFileName As String

Public Function SetErrorLog(nErrorNo As Long, strErDes As String, strMetName As String, sFileName As String)
    'Set all data members
    SetErrNumber (nErrorNo)
    SetErrDesc (strErDes)
    SetMethodName (strMetName)
    SetFileName (sFileName)
    'Write information to log file
    Call WriteToFile
End Function

Private Function SetErrNumber(nErrorNo As Long)
    nErrorNo = nErrorNo
End Function

Private Function SetErrDesc(str As String)
    strErrDesc = str
End Function
Private Function SetMethodName(str As String)
    strMethodName = str
End Function

Private Function SetFileName(str As String)
    strFileName = str
End Function


Private Sub cmdSendErrorReport_Click()
 'Another way to send this report is via web page.
 'For e.g. www.yourcompanyname.com/submit.aspx?...
 'This page will read the text and update database.. and corresponding user or developer
 'is assigned that defect...

 'Dim Handle As Integer
 'Handle = FreeFile
  
 'Pass first parameter to "false" if you do not want to show outlook send dialog to user.
 SendMessage True, App.Path & "\err.log"
  
 'You can check whether message was send successfully or not.
 'Clear log file
 'Pending clear log functionality
 'Open App.Path & "\err.log" For Output As #Handle
 'Close #Handle
 
 'Terminate application
 
 MsgBox "Thanks for sending error log." & vbCrLf & "This will help us to improve product quality (please modify this msgbox )", vbOKOnly
 End
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblDescript.Caption = "We have created an error report that you can send to help us improve" & vbCrLf & _
                          "Application Name. We will treat this report as confidential and " & vbCrLf & "anonymous."
                          
    'Initialize all members to null
    SetErrNumber (0)
    SetErrDesc ("")
    SetMethodName ("")
    SetFileName ("")
End Sub

Private Sub Label5_Click()
    Dim strPath As String
    strPath = App.Path & "\err.log"
    Call ShellExecute(0, "open", strPath, 0, 0, 1)
End Sub

'You can modify below code to customize as per your need.
'for e.g. if you need any dll version information. etc.
'any machine information then that can be found out.
'O.S. Version....

Private Function WriteToFile()
On Error GoTo ErrorHandler
    Dim Handle As Integer
    Handle = FreeFile
    
    Open App.Path & "\err.log" For Append As #Handle
    
    Print #Handle, "***********************************************************"
    Print #Handle, "Error Info"
    Print #Handle, "File Name                 :" & strFileName
    Print #Handle, "Method Name               :" & strMethodName
    Print #Handle, "Error Description         :" & strErrDesc
    Print #Handle, "Error Number              :" & nErrNo & vbCrLf
    Print #Handle, "System Information        :"
    Print #Handle, "OS Information            :" & GetOSVersion
    Print #Handle, "CPU Type                  :" & GetCPUType
    Print #Handle, "Current Time              :" & GetLocalDateTime
    
    'For E.g. to get directX version below code can be used.
    'Print #Handle, "DirectX Version           :" & GetDirectXVersion()
    'just print or update here with your data...
    Print #Handle, "***********************************************************"
    On Error Resume Next
    Close #Handle
    Exit Function
ErrorHandler:
    On Error Resume Next
    Close #Handle
End Function


Private Sub XPButton1_Click()
 'Another way to send this report is via web page.
 'For e.g. www.yourcompanyname.com/submit.aspx?...
 'This page will read the text and update database.. and corresponding user or developer
 'is assigned that defect...

 'Dim Handle As Integer
 'Handle = FreeFile
  
 'Pass first parameter to "false" if you do not want to show outlook send dialog to user.
 SendMessage True, App.Path & "\err.log"
  
 'You can check whether message was send successfully or not.
 'Clear log file
 'Pending clear log functionality
 'Open App.Path & "\err.log" For Output As #Handle
 'Close #Handle
 
 'Terminate application
 
 MsgBox "Thanks for sending error log." & vbCrLf & "This will help us to improve product quality (please modify this msgbox )", vbOKOnly
 End
End Sub

Private Sub XPButton2_Click()
    Unload Me
End Sub
