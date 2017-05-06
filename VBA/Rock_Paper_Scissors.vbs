Option Explicit

Public Sub main()
	Dim n As Integer
	Dim ia_Score As Integer
	Dim human_Score As Integer


	n = InputBox("How many moves to beat the IA")

	If (n = 0) Then
		MsgBox ("Invalid entry moves has to be superior as 1 ")
	End If

	While (human_Score <> n And ia_Score <> n)

		Call Compare(human_Move, ia_Move(random()), ia_Score, human_Score)

	Wend

	If (human_Score = n) Then
		MsgBox ("You WIN" & human_Score & "to" & ia_Score)
	ElseIf (ia_Score = n) Then
		MsgBox ("You Loose" & human_Score & "to" & ia_Score)
	End If


End Sub
Function random() As Integer

	Randomize
	random = Int((3 - 1 + 1) * Rnd + 1)



End Function

Function ia_Move(random_number As Integer) As String

	If random_number = 1 Then
		ia_Move = "rock"
	ElseIf random_number = 2 Then
		ia_Move = "paper"
	Else
		ia_Move = "scissor"

	End If
End Function

Function human_Move() As String
	Dim move As Integer

	While move <> 1 And move <> 2 And move <> 3

		move = InputBox(" Choose your move" & vbNewLine & " 1 For rock" & vbNewLine & " 2 For paper " & vbNewLine & " 3 For Scissor")
		If move = 1 Then
			human_Move = "rock"
		ElseIf move = 2 Then
			human_Move = "paper"
		ElseIf move = 3 Then
			human_Move = "scissor"
		Else
			MsgBox ("invalid Choice")

		End If
	Wend

End Function

Private Sub Compare(human_Move As String, ia_Move As String, ia_Score As Integer, human_Score As Integer)


	If human_Move = ia_Move Then
		MsgBox ("Draw")
	ElseIf (human_Move = "rock" And ia_Move = "paper") Then
		ia_Score = ia_Score + 1
		MsgBox ("Ia WIN")
	ElseIf (human_Move = "scissor" And ia_Move = "rock") Then
		ia_Score = ia_Score + 1
		MsgBox ("Ia WIN")
	ElseIf (human_Move = "rock" And ia_Move = "scissor") Then
		human_Score = human_Score + 1
		MsgBox ("You WIN")
	ElseIf (human_Move = "scissor" And ia_Move = "paper") Then
		human_Score = human_Score + 1
		MsgBox ("You WIN")
	ElseIf (human_Move = "paper" And ia_Move = "scissor") Then
		ia_Score = ia_Score + 1
		MsgBox ("Ia WIN")
	ElseIf (human_Move = "paper" And ia_Move = "rock") Then
		human_Score = human_Score + 1
		MsgBox ("You WIN")

	End If
End Sub
