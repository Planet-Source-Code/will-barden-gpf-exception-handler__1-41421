VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Never get caught by another GPF again!"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top GPF Handler"
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start GPF Handler"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdError 
      Caption         =   "&Raise GPF"
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UN-SAFE!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdError_Click()
Dim lRet As Long
Dim s As String
Dim b(5) As Byte
On Error GoTo Whoops
   TestGPFHandler
   Exit Sub
Whoops:
   ' call graceful shutdown code here
   With frmException
      .Label2.Caption = Err.Description
      .Show vbModal
   End With
   If Not frmException.Continue Then
      Unload frmMain
      Unload frmException
   End If
End Sub

Private Sub cmdStart_Click()
   
   ' start the safeguard
   If StartGPFHandler Then
      lblStatus.ForeColor = vbBlue
      lblStatus.Caption = "SAFE"
   Else
      lblStatus.ForeColor = vbRed
      lblStatus.Caption = "FAILED TO START: UN-SAFE"
   End If
   
End Sub

Private Sub cmdStop_Click()

   ' stop handling..
   StopGPFHandler
   lblStatus.ForeColor = vbRed
   lblStatus.Caption = "UN-SAFE"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StopGPFHandler
End Sub
