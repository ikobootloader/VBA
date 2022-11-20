Sub repartition()
    
    '/** REPARTITION DOSSIERS PAR ORDRE ALPAHABETIQUE **/
    
    'Définir la quantité de données (noms)
    Dim quantiteDonnees As Long
    quantiteDonnees = Range("A1:A210").Count

    'Définir la quantité d'instructrices
    Dim nbInstruct As Integer
    nbInstruct = Range("C2").Value

    'Diviser cette quantité par le nombre d'instructriceµ
    Dim divDonneesParInstruct As Long
    divDonneesParInstruct = quantiteDonnees / nbInstruct
    Range("D2").Value = divDonneesParInstruct
    
    'Sur x plage : intégrer arrangements alphabétiques
    Dim alphabet() As Variant
    alphabet = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    
    'Compter le nombre de nom
    Dim compterDonnees As Long
    'Compter le nombre de nom
    Dim compterDonnees2 As Long
    'Compter le nombre de nom
    Dim compterDonnees3 As Long
    'Compter le nombre de nom
    Dim compterDonnees4 As Long
    
    'Compter jusqu'à 30
    Dim resultatCorresp As Long
    resultatCorresp = 0
    
    'Variable lettres
    Dim numeroCell As Integer
    numeroCell = 16
    
    'Tableau de répartition
    Dim lettres As Variant
    Dim repartition As Integer
    repartition = 2
    
    'Boucle
    For compterDonnees = 0 To 25
        
        'Intégrer les arrangements alphabétiques
        Range("F" & numeroCell & "").Value = alphabet(compterDonnees)
        'Intégrer les quantités de noms associés
        Range("G" & numeroCell & "").Select
        ActiveCell.FormulaLocal = "=NB.SI(A1:A210;""" & alphabet(compterDonnees) & "*"")"
        'MsgBox compterDonnees
        numeroCell = numeroCell + 1
    
    Next
    
    'Reset numeroCell
    numeroCell = 16
        
    For compterDonnees = 0 To 25
        
        For compterDonnees2 = 0 To 25
        
        'Intégrer les arrangements alphabétiques
        Range("H" & numeroCell & "").Value = alphabet(compterDonnees) + alphabet(compterDonnees2)
        'Intégrer les quantités de noms associés
        Range("I" & numeroCell & "").Select
        ActiveCell.FormulaLocal = "=NB.SI(A1:A210;""" & alphabet(compterDonnees) + alphabet(compterDonnees2) & "*"")"
        numeroCell = numeroCell + 1
        
        Next
        
        numeroCell = numeroCell + 1
                    
    Next
    
    'Reset numeroCell
    numeroCell = 16
    
    '***********************
    
    'Reset numeroCell
    numeroCell = 219
    
    'Variable colonne
    Dim colonne As Integer
    colonne = 0

    Dim compteur As Long
    compteur = 0
    
    Dim objet As Variant
    objet = 0
    
    Dim lettreA As Variant
    lettreA = ""
    
    Dim bool As Integer
    bool = 0
    
    For compterDonnees = 0 To 25
    
        For compterDonnees2 = 0 To 25
        
            'Reset si on atteint AA
             
            For compterDonnees3 = 0 To 25
                
                'If colonne = 25 Then
                'MsgBox compterDonnees3
                'End If
                
                'Intégrer les arrangements alphabétiques
                Range("" & lettreA & alphabet(colonne) & numeroCell & "").Value = alphabet(compterDonnees) + alphabet(compterDonnees2) + alphabet(compterDonnees3)
                'Intégrer les quantités de noms associés
                Range("" & lettreA & alphabet(colonne + 1) & numeroCell & "").Select
                ActiveCell.FormulaLocal = "=NB.SI(A1:A210;""" & alphabet(compterDonnees) + alphabet(compterDonnees2) + alphabet(compterDonnees3) & "*"")"
                numeroCell = numeroCell + 1

                objet = lettreA & alphabet(colonne)
                If objet = "Y" Then
                    If numeroCell = "245" Then
                        'MsgBox "limite"
                        colonne = 0
                        lettreA = "A"
                    End If
                    If numeroCell = 245 + 27 * compteur Then
                        'MsgBox "limite"
                        colonne = 0
                        lettreA = "A"
                    End If
                End If
                
            Next
           
            numeroCell = 219
            If compteur > 0 Then
                numeroCell = 219 + 27 * compteur
            End If
            
            'Colonnes préfixées A
            objet = lettreA & alphabet(colonne)
            'MsgBox objet
            If objet = "AA" Then
                If bool <> 1 Then
                    'MsgBox "AA"
                    colonne = -2
                    bool = 1
                End If
            End If
            
            'On part sur les colonnes adjacentes
            colonne = colonne + 2
             
        Next
   
            'Reset pour toutes les autres lettres
            compteur = compteur + 1
            numeroCell = 219 + 27 * compteur
            colonne = 0
            lettreA = ""
            bool = 0
            
    Next
    'Si quantité n'est pas exact à la quantité de dossier théorique par instructrice (divDonneesParInstruct)
    'Stopper boucle et Ajouter un arrangement supplémentaire
    'Insérer les arrangements alpahbétiques limites pour chaque instructrice : de x lettre(s) à y lettre(s)

End Sub
