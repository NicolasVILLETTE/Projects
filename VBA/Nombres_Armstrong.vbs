Option Explicit

Dim Nombre As Integer
Dim Centaine As Integer
Dim Dizaine As Integer
Dim Unité As Integer

Sub main()

	Nb_Armstrong

End Sub

Public Function Nb_Armstrong()

	Nombre = InputBox("Veuillez entrer un nombre entre 100 et 499:")

	If Nombre < 100 Or Nombre > 499 Then
		Nombre = InputBox("Erreur: Veuillez entrer un nombre entre 100 et 499:")
	Else
		Call Decomposition_Nombre(Nombre)
		If Is_Armstrong(Nombre, Centaine, Dizaine, Unité) = True Then
			MsgBox ("Le nombre" & Nombre & " est un nombre d'Armstrong")
		Else
			MsgBox ("Le nombre" & Nombre & " n'est pas un nombre d'Armstrong")
		End If
	End If

End Function

Public Function Decomposition_Nombre(Nombre)

	Centaine = Int(Nombre / 100)
	Dizaine = Int((Nombre - Centaine * 100) / 10)
	Unité = Int(Nombre - (Centaine * 100 + Dizaine * 10))

	MsgBox (Centaine)
	MsgBox (Dizaine)
	MsgBox (Unité)


End Function

Public Function Is_Armstrong(ByRef Nombre, ByRef Centaine, ByRef Dizaine, ByRef Unité) As Boolean

	Dim somme_carrés As Integer

	somme_carrés = Centaine ^ 3 + Dizaine ^ 3 + Unité ^ 3

	MsgBox (somme_carrés)


	If Nombre = somme_carrés Then
		Is_Armstrong = True
	End If

End Function
