VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScoreSubmissionUserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6945
   OleObjectBlob   =   "ScoreSubmissionUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScoreSubmissionUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub OpenDoc()
    ScoreSubmissionUserForm.Show
End Sub

Private Sub CommandButton1_Click()
    ScoreSubmissionUserForm.Show
End Sub

Private Sub EndSubmission_Click()
    Unload Me
End Sub


Private Sub NextScore_Click()
        Dim scoreCol As Integer
            'Searches for the column to edit
            With Range("B1:Y1")
                Set rFind = .Find(What:=RecitationNumber.Value, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
                If Not rFind Is Nothing Then
                    scoreCol = rFind.Column
                End If
                If scoreCol = 0 Then
                        MsgBox RecitationNumber.Value
                End If
            End With
            
            'Finds and updates rows based on selected students
            With Range("A3:A400")
                Set rFind = .Find(What:=NameBox1.Value, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
                If Not rFind Is Nothing And NameBox1.Value <> "" Then
                    Cells(rFind.Row, scoreCol).Value = Score.Value
                End If
            End With
            With Range("A3:A400")
                Set rFind = .Find(What:=NameBox2.Value, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
                If Not rFind Is Nothing And NameBox2.Value <> "" Then
                    Cells(rFind.Row, scoreCol).Value = Score.Value
                End If
            End With
            With Range("A3:A400")
                Set rFind = .Find(What:=NameBox3.Value, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
                If Not rFind Is Nothing And NameBox3.Value <> "" Then
                    Cells(rFind.Row, scoreCol).Value = Score.Value
                End If
            End With
            With Range("A3:A400")
                Set rFind = .Find(What:=NameBox4.Value, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
                If Not rFind Is Nothing And NameBox4.Value <> "" Then
                     Cells(rFind.Row, scoreCol).Value = Score.Value
                End If
            End With
            
            If RecitationNumber.Value <> "" Then
             ActiveColumn = RecitationNumber.Value
            End If
            
            
   Unload Me
   ScoreSubmissionUserForm.Show
   
   
   
   
End Sub

Private Sub UserForm_Initialize()
    Dim ActiveColumnNumber As Integer
    Dim MyCell As Variant
    'Get Names from first Column
    For Each MyCell In Range("A3:A300")
        If MyCell.Value = "" Then
            Exit For
        End If
        With NameBox1
            .AddItem MyCell.Value
        End With
        With NameBox2
            .AddItem MyCell.Value
        End With
        With NameBox3
            .AddItem MyCell.Value
        End With
        With NameBox4
            .AddItem MyCell.Value
        End With
    
        'Stops Searching when it reaches Student Test
        If (StrComp(MyCell.Value, "Student, Test", vbTextCompare) = 0) Then
            Exit For
        End If
    
    Next MyCell
    
    'Gets Column Names
    For Each MyCell In Range("D1:BZ1")
        If MyCell.Value = "" Then
            Exit For
        End If
        'Selects the active Column
        If ActiveColumn = MyCell.Value Then
            RecitationNumber.AddItem MyCell.Value
            ActiveColumnNumber = RecitationNumber.ListCount - 1
        Else
            RecitationNumber.AddItem MyCell.Value
        End If
    Next MyCell
    RecitationNumber.Selected(ActiveColumnNumber) = True
End Sub

