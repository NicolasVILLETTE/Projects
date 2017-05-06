Option Explicit

Sub main()

	Dim borne As Integer
	borne = InputBox("Entrer un entier comme borne maximale de recherche:")
	Est_Parfait (borne)

End Sub



Function Est_Parfait(borne) As Integer

	Dim Nombre As Integer
	Dim Somme As Integer
	Dim i As Integer

	For Nombre = 0 To borne
		Somme = 0
		For i = 1 To Nombre - 1
			If Nombre Mod i = 0 Then
				Somme = Somme + i
			End If
			Next i

			If Somme = i Then
				MsgBox "Le nombre " & Nombre & " est un nombre parfait"
			End If
			Next Nombre
		End Function
