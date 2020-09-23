VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "VBScripting .: By Justin Lilley"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "&End"
      Height          =   735
      Left            =   5520
      TabIndex        =   10
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Script"
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmmain.frx":0000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open Script"
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Text            =   "Notes:"
      Top             =   5520
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reload"
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox pic1 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton Command9 
         Caption         =   "P4"
         Height          =   375
         Left            =   4080
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   5295
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   4080
         Top             =   360
      End
      Begin VB.CommandButton Command7 
         Caption         =   "D3"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "A0"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "C2"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3480
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "B1"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4935
      End
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   4200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   3375
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'.: Vbscripting by Justin Lilley aka 715
'.: Just whiped this up tonite - 12/17/04 12:22 PM
'.: But if you do like it please vote.
'.: Of leave Feedback on this code.

Private Sub Command1_Click()
ScriptControl1.ExecuteStatement "B1_click"
End Sub
Private Sub Command2_Click()
ScriptControl1.AddCode RTB.Text
End Sub

Private Sub Command3_Click()
    RTB.LoadFile (App.Path & "\Script.txt")
    Me.Caption = "VBScripting .: By Justin Lilley"
End Sub

Private Sub Command4_Click()
    CommonDialog1.Filter = "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
    CommonDialog1.ShowSave
    RTB.SaveFile (CommonDialog1.FileName)
    Me.Caption = "VBScripting .: By Justin Lilley"
End Sub

Private Sub Command5_Click()
ScriptControl1.ExecuteStatement "c2_click"
End Sub

Private Sub Command6_Click()
ScriptControl1.ExecuteStatement "a0_click"
End Sub

Private Sub Command7_Click()
ScriptControl1.ExecuteStatement "d3_click"
End Sub

Private Sub Command8_Click()
End
End Sub

Private Sub Command9_Click()
ScriptControl1.ExecuteStatement "p4_click"
End Sub

Private Sub Form_Load()
RTB.LoadFile (App.Path & "\Script.txt")
ScriptControl1.AddObject "Tim1", Timer1
ScriptControl1.AddObject "Tim2", Timer2
ScriptControl1.AddObject "form1", pic1
ScriptControl1.AddObject "form", Form1
ScriptControl1.AddObject "b1", Command1
ScriptControl1.AddObject "c2", Command5
ScriptControl1.AddObject "d3", Command7
ScriptControl1.AddObject "p4", Command9
ScriptControl1.AddObject "lbl1", Label1
ScriptControl1.AddObject "chk1", Check1
ScriptControl1.AddCode RTB.Text
ScriptControl1.Run "Main"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ScriptControl1.Run "Main"
ScriptControl1.ExecuteStatement "Tim1_Timer"
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ScriptControl1.Run "Main"
ScriptControl1.ExecuteStatement "Tim2_Timer"
End Sub
