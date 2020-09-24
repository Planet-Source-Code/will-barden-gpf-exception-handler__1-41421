VERSION 5.00
Begin VB.Form frmException 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Whoops!"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&Continue"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   $"frmException.frx":0000
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "An unrecoverable error has occurred... oh no!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -------------------------------------------------------------- '
' form to be shown when an unhandled exception (GPF) occurs
' created 25/11/02
' modified 25/11/02
' will barden
' -------------------------------------------------------------- '

' -------------------------------------------------------------- '
' private variables
' -------------------------------------------------------------- '

Private mbContinue As Boolean

' -------------------------------------------------------------- '
' extra properties
' -------------------------------------------------------------- '

Public Property Get Continue() As Boolean
   Continue = mbContinue
End Property

' -------------------------------------------------------------- '
' form events
' -------------------------------------------------------------- '

Private Sub cmdContinue_Click()

   ' set the flag
   mbContinue = True

   ' remove the form
   Me.Hide

End Sub

Private Sub cmdExit_Click()

   ' set the flag
   mbContinue = False

   ' remove the form
   Me.Hide

End Sub

