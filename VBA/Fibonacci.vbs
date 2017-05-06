Option Explicit
Public Sub main()
	Dim n As Integer

	n = InputBox("Wich term do you want to know ?")
	Fibonacci (n)


End Sub

Function Fibonacci(n) As Double
	Dim result As Double
	Dim first As Double
	Dim second As Double
	Dim k As Double

	result = 0
	first = 0
	second = 1

	If n = 0 Then
		result = first
	ElseIf n = 1 Then
		result = second
	Else
		For k = 2 To n
			result = first + second
			first = second
			second = result
			Next k

		End If

		MsgBox (" The " & n & "term of the fibonacci sequence is " & result)

	End Function
