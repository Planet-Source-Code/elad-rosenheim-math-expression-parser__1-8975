VERSION 5.00
Begin VB.Form frmParserDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Math Expression Parser"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   " Parser "
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtExpression 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Text            =   "1+1"
         Top             =   480
         Width           =   2895
      End
      Begin VB.CommandButton cmdParse 
         Caption         =   "Parse"
         Height          =   375
         Left            =   3120
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdMultiParse 
         Caption         =   "Multi-Parse!"
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtNumParses 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Text            =   "1000"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Expression:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Times to parse:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1000
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Constants "
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   4575
      Begin VB.CommandButton cmdAddConst 
         Caption         =   "Add Constant"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtConstName 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtConstValue 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.ListBox lstConstants 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton cmdRemoveConst 
         Caption         =   "Remove Const"
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Constant name:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Value:"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " More Info "
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   4575
      Begin VB.Label Label5 
         Caption         =   $"Form1.frx":08CA
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   310
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmParserDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyParser As New clsExpressionParser

' Keep track of user-defined constants
Dim ConstNames As New Collection

Private Sub cmdAddConst_Click()
On Error GoTo cmdAddConst_ErrHandler
    
    ' Validity checks
    If Trim(txtConstName) = "" Then
        MsgBox "A valid constant name must begin with a letter. " & _
            "Additional letters may include the alphanumeric letters or underscores. " & _
            "For Example: MyConst_2", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(txtConstValue) Then
        MsgBox "Please enter a valid number!", vbExclamation
        Exit Sub
    End If
    
    MyParser.AddConstant txtConstName, CDbl(txtConstValue)
    
    lstConstants.AddItem txtConstName & " - " & txtConstValue
    ConstNames.Add txtConstName.Text
    
    Exit Sub

cmdAddConst_ErrHandler:
    ' If the AddConstant call raises an error, it's
    ' description will be shown to the user
    MsgBox "Error:" & vbCrLf & Err.Description, _
            vbCritical
End Sub

Private Sub cmdParse_Click()
On Error GoTo cmdParse_ErrHandler

Dim Result As Double

    Result = MyParser.ParseExpression(txtExpression.Text)
    MsgBox "Result:  " & Format(Result, "#0.0#####")
    Exit Sub
    
cmdParse_ErrHandler:
    ' If the error was raised intentionally by the parser
    ' (becuase of some syntax error), show a detailed
    ' message
    If Err.Number >= PERR_FIRST And _
       Err.Number <= PERR_LAST Then
        ShowParseError
    Else
        MsgBox Err.Description, vbCritical, "Unexpected Error"
    End If
End Sub

Private Sub cmdMultiParse_Click()
On Error GoTo cmdMultiParse_ErrHandler

Dim Value As Integer
Dim i As Long
Dim Expression As String
Dim StartTime As Single, EndTime As Single
Dim NumParses As Long

    If Not IsNumeric(txtNumParses) Or Val(txtNumParses) < 1 Then
        MsgBox "Please enter a positive number!", vbExclamation
        Exit Sub
    End If
    
    NumParses = CLng(txtNumParses)
    Expression = txtExpression.Text
    
    StartTime = Timer
    For i = 1 To NumParses
        Value = MyParser.ParseExpression(Expression)
    Next
    EndTime = Timer

    MsgBox "Time took:" & vbCrLf & _
        Format(CStr(EndTime - StartTime), "#0.0##") & " Seconds", _
        vbInformation
    Exit Sub
    
cmdMultiParse_ErrHandler:
    If Err.Number >= PERR_FIRST And _
       Err.Number <= PERR_LAST Then
        ShowParseError
    Else
        MsgBox Err.Description, vbCritical, "Unexpected Error"
    End If
End Sub

Private Sub ShowParseError()
    
    ' Show error details
    MsgBox "Error No. " & CStr(Err.Number - PERR_FIRST + 1) & _
        " - " & Err.Description & vbCrLf & _
        "Raised from: " & Err.Source, _
        vbCritical, "Parse Error"
    
    ' Mark the position in the expression where the error
    ' was raised
    txtExpression.SelStart = MyParser.LastErrorPosition - 1
    txtExpression.SelLength = 1
    txtExpression.SetFocus

End Sub

Private Sub cmdRemoveConst_Click()
On Error GoTo cmdRemoveConst_ErrHandler

    If lstConstants.ListIndex = -1 Then
        MsgBox "Select a constant to remove!", vbExclamation
        Exit Sub
    End If

    MyParser.RemoveConstant ConstNames(lstConstants.ListIndex + 1)
    
    ConstNames.Remove lstConstants.ListIndex + 1
    lstConstants.RemoveItem lstConstants.ListIndex
    Exit Sub
    
cmdRemoveConst_ErrHandler:
    MsgBox "Error:" & vbCrLf & Err.Description, _
            vbCritical
End Sub
