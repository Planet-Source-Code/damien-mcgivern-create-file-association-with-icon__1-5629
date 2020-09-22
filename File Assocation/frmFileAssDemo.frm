VERSION 5.00
Begin VB.Form frmFileAssDemo 
   Caption         =   "File Association Demo"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   Icon            =   "frmFileAssDemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkApp 
      Caption         =   "Start with this application"
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Left            =   1440
      TabIndex        =   18
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CheckBox chkDefault 
      Caption         =   "Use Application's Icons"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtIcon 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   2640
      Width           =   6375
   End
   Begin VB.CheckBox chkIcon 
      Caption         =   "Change Icon"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtTypeName 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "My File Type's Name"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtType 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "My.FileType"
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtExt 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   ".extension"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "Open Me Up"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "The number at the end corrisponds to the resorce file's icons 1 = 101, 2 = 102 etc"
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Label Label9 
      Caption         =   "ie C:\program files\My prog.exe"
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Application"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "ie Visual Basic Form"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "ie VB.Form"
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Must start with a full stop ie .zip"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "ie Open"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblIcon 
      Caption         =   "Icon"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "File Type Name"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "File Type"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lblExt 
      Caption         =   "Extension "
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblComand 
      Caption         =   "Command "
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmFileAssDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkApp_Click()

    If Not chkApp.Value = 1 Then
        txtApp.Text = ""
    Else
        txtApp.Text = App.Path & "\" & App.EXEName & ".exe"
    End If
    
End Sub

Private Sub chkDefault_Click()
    If Me.chkDefault.Value = 1 Then Me.txtIcon.Text = App.Path & "\" & App.EXEName & ".exe,1"
End Sub

Private Sub chkIcon_Click()
    
    
    If Me.chkIcon.Value = 1 Then
        Me.chkDefault.Enabled = 1
        Me.txtIcon.Enabled = Not chkDefault.Value
        Me.lblIcon.Enabled = Not chkDefault.Value
    Else
        Me.chkDefault.Enabled = 0
        Me.txtIcon.Enabled = False
        Me.lblIcon.Enabled = False
    End If
End Sub

Private Sub cmdApply_Click()
    If Me.txtCommand.Text = "" Then
        MsgBox "Enter a command"
        Me.txtCommand.SetFocus
        
    ElseIf Me.txtApp.Text = "" Then
        MsgBox "Enter associated application's path"
        Me.txtApp.SetFocus
        
    ElseIf Me.txtExt.Text = "" Then
        MsgBox "Enter a file extension"
        Me.txtExt.SetFocus
        
    ElseIf Not Mid(Me.txtExt.Text, 1, 1) = "." Then
        MsgBox "Extension must start with '.'"
        Me.txtExt.SetFocus
        
    ElseIf Me.txtType.Text = "" Then
        MsgBox "Enter a file type"
        Me.txtType.SetFocus
        
    ElseIf Me.txtTypeName.Text = "" Then
        MsgBox "Enter a file type name"
        Me.txtTypeName.SetFocus
        
    ElseIf Me.chkIcon.TabIndex = 1 And Me.txtIcon = "" Then
        MsgBox "Enter a path for the default icon"
        Me.chkIcon.SetFocus
        
    Else
    
        
        
        If FileAss.CreateFileAss(txtExt, txtType, txtTypeName, txtCommand, txtApp, chkIcon, txtIcon) Then
            MsgBox "File assocation for '" & txtExt & "' created"
        Else
            MsgBox "An error occured while assocating the extinsion '" & txtExt & "'"
        End If
        
    End If
    
    
End Sub

Private Sub Form_Load()
    If Not Command = "" Then MsgBox " I was opened with the file '" & Command & "'"
End Sub
