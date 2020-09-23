VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TicTacToe"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Row3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Tag             =   "9"
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Row3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   840
      TabIndex        =   7
      Tag             =   "8"
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Row3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Tag             =   "7"
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Row2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   1680
      TabIndex        =   5
      Tag             =   "6"
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Row2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   840
      TabIndex        =   4
      Tag             =   "5"
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Row2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Tag             =   "4"
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Row1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Tag             =   "3"
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Row1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Tag             =   "2"
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Row1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Tag             =   "1"
      Top             =   0
      Width           =   795
   End
   Begin VB.Line Line4 
      X1              =   1680
      X2              =   1680
      Y1              =   0
      Y2              =   1800
   End
   Begin VB.Line Line3 
      X1              =   840
      X2              =   840
      Y1              =   0
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   2520
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2520
      Y1              =   555
      Y2              =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr(1 To 3, 1 To 3) As Integer
Dim lastID As Integer
Dim firstMove As Boolean
Dim secondMove As Boolean

Private Sub Form_Load()
MsgBox "Please Make necessaey Note" & vbCrLf & "Author Name: Syed Rizwan Muhammad Rizvi" & vbCrLf & "Roll No: 268" & vbCrLf & "MCS-R-II" & vbCrLf & "AI Assignment No.2"
resetControls
End Sub

Private Sub Row1_Click(Index As Integer)
If Row1(Index).Caption = "" Then
    Row1(Index).Caption = "X"
Else
    Exit Sub
End If
'lastID = Index + 1
lastID = getID(0, Index)
arr(1, Index + 1) = 1
doGoalTest
makeMove
End Sub
Private Sub Row2_Click(Index As Integer)
If Row2(Index).Caption = "" Then
    Row2(Index).Caption = "X"
Else
    Exit Sub
End If
'lastID = Index + 1
lastID = getID(1, Index)
arr(2, Index + 1) = 1
doGoalTest
makeMove
End Sub
Private Sub Row3_Click(Index As Integer)
If Row3(Index).Caption = "" Then
    Row3(Index).Caption = "X"
Else
    Exit Sub
End If
'lastID = Index + 1
lastID = getID(2, Index)
arr(3, Index + 1) = 1
doGoalTest
makeMove
End Sub

Function evalGoal() As Integer
evalGoal = -100
Dim sm As Integer
Dim zCnt As Integer
Dim k As Integer
Dim i As Integer
'Check Horizontal
zCnt = 0
For k = 1 To 3
    sm = 0
    For i = 1 To 3
        sm = sm + arr(k, i)
        If arr(k, i) = 0 Then zCnt = zCnt + 1
    Next
    If Abs(sm) = 3 Then
        evalGoal = sm
        Exit Function
    End If
Next
'Check Vertical
For k = 1 To 3
    sm = 0
    For i = 1 To 3
        sm = sm + arr(i, k)
    Next
    If Abs(sm) = 3 Then
        evalGoal = sm
        Exit Function
    End If
Next

'Check Diognal Left-->Right
sm = 0
For i = 1 To 3
    sm = sm + arr(i, i)
Next
If Abs(sm) = 3 Then
    evalGoal = sm
    Exit Function
End If

'Check Diognal Right-->Left
sm = 0
For i = 3 To 1 Step -1
    sm = sm + arr(4 - i, i)
Next
If Abs(sm) = 3 Then
    evalGoal = sm
    Exit Function
End If
If zCnt = 0 Then evalGoal = -500
End Function

Sub doGoalTest()
Dim k As Integer
k = evalGoal
If k = -500 Then
    MsgBox "Draw! You just Can't Win"
    resetControls
ElseIf Not k = -100 Then
    If k < 0 Then
        MsgBox "Computer Wins"
    Else
        MsgBox "You Win and if u do then this program is buggy"
    End If
    resetControls
End If
End Sub
Sub resetControls()
Dim i As Integer
Dim k As Integer
    For i = 1 To 3
        For k = 1 To 3
            arr(i, k) = 0
            Select Case i
                Case 1:
                    Row1(k - 1).Caption = ""
                Case 2:
                    Row2(k - 1).Caption = ""
                Case 3:
                    Row3(k - 1).Caption = ""
            End Select
        Next
    Next
lastID = -1
firstMove = True
secondMove = False
End Sub

Sub makeMove()
Dim sm As Integer
Dim k As Integer
Dim i As Integer
If lastID = -1 Then Exit Sub
    'Check for Win
    'Check Horizontal
    For k = 1 To 3
        sm = 0
        For i = 1 To 3
            sm = sm + arr(k, i)
        Next
        If sm = -2 Then
            For i = 1 To 3
                If arr(k, i) = 0 Then
                    arr(k, i) = -1
                    setLabel k - 1, i - 1, "O"
                    GoTo ed:
                End If
            Next
        End If
    Next

    'Check Vertical
    For k = 1 To 3
        sm = 0
        For i = 1 To 3
            sm = sm + arr(i, k)
        Next
        If sm = -2 Then
            For i = 1 To 3
                If arr(i, k) = 0 Then
                    arr(i, k) = -1
                    setLabel i - 1, k - 1, "O"
                    GoTo ed:
                End If
            Next
        End If
    Next

    'Check Diognal Left-->Right
    sm = 0
    For i = 1 To 3
        sm = sm + arr(i, i)
    Next
    If sm = -2 Then
            For i = 1 To 3
                If arr(i, i) = 0 Then
                    arr(i, i) = -1
                    setLabel i - 1, i - 1, "O"
                    GoTo ed:
                End If
            Next
        End If

    'Check Diognal Right-->Left
    sm = 0
    For i = 3 To 1 Step -1
        sm = sm + arr(4 - i, i)
    Next
    If sm = -2 Then
            For i = 3 To 1 Step -1
                If arr(4 - i, i) = 0 Then
                    arr(4 - i, i) = -1
                    setLabel 4 - i - 1, i - 1, "O"
                    GoTo ed:
                End If
            Next
    End If

    'Check for Loss
    'Check Horizontal
    For k = 1 To 3
        sm = 0
        For i = 1 To 3
            sm = sm + arr(k, i)
        Next
        If Abs(sm) = 2 Then
            For i = 1 To 3
                If arr(k, i) = 0 Then
                    arr(k, i) = -1
                    setLabel k - 1, i - 1, "O"
                    GoTo ed:
                End If
            Next
        End If
    Next

    'Check Vertical
    For k = 1 To 3
        sm = 0
        For i = 1 To 3
            sm = sm + arr(i, k)
        Next
        If Abs(sm) = 2 Then
            For i = 1 To 3
                If arr(i, k) = 0 Then
                    arr(i, k) = -1
                    setLabel i - 1, k - 1, "O"
                    GoTo ed:
                End If
            Next
        End If
    Next

    'Check Diognal Left-->Right
    sm = 0
    For i = 1 To 3
        sm = sm + arr(i, i)
    Next
    If Abs(sm) = 2 Then
            For i = 1 To 3
                If arr(i, i) = 0 Then
                    arr(i, i) = -1
                    setLabel i - 1, i - 1, "O"
                    GoTo ed:
                End If
            Next
        End If

    'Check Diognal Right-->Left
    sm = 0
    For i = 3 To 1 Step -1
        sm = sm + arr(4 - i, i)
    Next
    If Abs(sm) = 2 Then
            For i = 3 To 1 Step -1
                If arr(4 - i, i) = 0 Then
                    arr(4 - i, i) = -1
                    setLabel 4 - i - 1, i - 1, "O"
                    GoTo ed:
                End If
            Next
    End If

If (Not (lastID Mod 2) = 0) And (Not (lastID = 5)) And firstMove = True Then
    Row2(1).Caption = "O"
    arr(2, 2) = -1
ElseIf lastID = 5 And firstMove Then
    arr(1, 1) = -1
    setLabel 1 - 1, 1 - 1, "O"
ElseIf lastID = 7 And secondMove And arr(2, 2) = 0 Then
    arr(2, 2) = -1
    setLabel 2 - 1, 2 - 1, "O"
ElseIf lastID = 8 And secondMove Then
    arr(3, 1) = -1
    setLabel 3 - 1, 1 - 1, "O"
ElseIf lastID = 9 And secondMove Then
    arr(1, 3) = -1
    setLabel 1 - 1, 3 - 1, "O"
ElseIf lastID = 6 And secondMove Then
    arr(3, 3) = -1
    setLabel 3 - 1, 3 - 1, "O"
Else
    For k = 1 To 3
        For i = 1 To 3
            If arr(k, i) = 0 Then
                If lastID Mod 2 = 0 And getID(k - 1, i - 1) Mod 2 = 1 Then
                    'Odd move
                    arr(k, i) = -1
                    setLabel k - 1, i - 1, "O"
                    GoTo ed:
                ElseIf lastID Mod 2 = 1 And getID(k - 1, i - 1) Mod 2 = 0 Then
                    'Even move
                    arr(k, i) = -1
                    setLabel k - 1, i - 1, "O"
                    GoTo ed:
                End If
            End If
        Next i
    Next k
    For k = 1 To 3
        For i = 1 To 3
            If arr(k, i) = 0 Then
                arr(k, i) = -1
                setLabel k - 1, i - 1, "O"
                GoTo ed:
            End If
        Next i
    Next k
End If
ed:
    If firstMove And Not secondMove Then
        firstMove = False
        secondMove = True
    Else
        secondMove = False
    End If
    doGoalTest
End Sub

Sub setLabel(rw As Integer, cl As Integer, tx As String)
Select Case rw
    Case 0:
        Row1(cl).Caption = tx
    Case 1:
        Row2(cl).Caption = tx
    Case 2:
        Row3(cl).Caption = tx
End Select
End Sub

Function getID(rw As Integer, cl As Integer) As Integer
Select Case rw
    Case 0:
        getID = CInt(Row1(cl).Tag)
    Case 1:
        getID = CInt(Row2(cl).Tag)
    Case 2:
        getID = CInt(Row3(cl).Tag)
End Select
End Function
