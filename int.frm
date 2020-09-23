VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Integration Machine"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   135
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton runEQ 
      Caption         =   "Calculate Function"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox delta 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Text            =   "0.00001"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox range2 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "1"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox range1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "0"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox comm 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "4*(sqr(1-x^2))"
      Top             =   360
      Width           =   3375
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   0
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Delta X:"
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Upper Bound:"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Lower Bound:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lbl1 
      Caption         =   "Formula f(x):"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +++++++++++++++++++++++++++++++++++++++++++++
' Demonstration of Advanced Mathematics in VB6
' Created with Chang Wui Meng and Roy Chai
' Royalty free distribution, as long as you
' Don't severe the codes :)
' THE INTEGRATOR 2003
' +++++++++++++++++++++++++++++++++++++++++++++
Private Sub Command1_Click()
End
End Sub

Private Sub runEQ_Click()
runEQ.Enabled = False
Dim r, a, d, z, y, q, s As Double
r = Int((range2.Text - range1.Text) / delta.Text)
a = range1.Text
d = delta.Text
sc.Reset
sc.AddCode "Function ExecStat(x)" & vbCrLf & "ExecStat=" & comm.Text & vbCrLf & "End Function"
For z = 1 To r - 1
pbar.Value = Round((z / r) * 100)
DoEvents
del = z * d
x = a + del
y = y + sc.Run("ExecStat", x)
Next
q = CInt(y * d * 10000) / 10000
comm.Text = "f(" & comm.Text & ") = " & q
sc.Reset
End Sub


Private Sub sc_Error()
sc.Reset
MsgBox "Please review your formula!", vbCritical, "Critical Error"
End Sub
