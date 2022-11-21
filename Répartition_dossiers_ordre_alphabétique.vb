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
    'MsgBox divDonneesParInstruct
    
    'Sur x plage : intégrer arrangements alphabétiques
    Dim alphabet() As Variant
    alphabet = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    
    'Compter le nombre de nom
    Dim compterDonnees, compterDonnees2, compterDonnees3, compterDonnees4 As Long
    
    'Compter jusqu'à 30
    'Dim resultatCorresp As Long
    'resultatCorresp = 0
    
    'Variable lettres
    Dim numeroCell As Integer
    numeroCell = 16
    
    'Tableau de répartition
    Dim lettres As Variant
    Dim repartition As Integer
    repartition = 2
    
    '***********************
    
    '/** ARRANGEMENTS ALAPHABETIQUE A 1 LETTRES **/
    
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
    
    '/** ARRANGEMENTS ALAPHABETIQUE A 2 LETTRES **/
    
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

    'Passage d'une lettre à l'autre dans la boucle
    Dim compteur As Long
    compteur = 0
    
    Dim objet As Variant
    objet = 0
    
    Dim lettreA As Variant
    lettreA = ""
    
    Dim bool As Integer
    bool = 0
    
    'Compter les noms jusqu'à 30
    Dim compterNoms As Integer
    compterNoms = 0
    
    'Colorisation des cellules
    Dim colorCell, colorCell2, colorCell3 As Integer
    colorCell = 30
    colorCell2 = 30
    colorCell3 = 50
    'colorCell = 255 / nbInstruct
    
    'Nom de la feuille
    Dim nomFeuille As Variant
    nomFeuille = ""
    'Compteur de feuille
    Dim compteurFeuille As Integer
    compteurFeuille = 0
    
    'Remplissage de cellules
    Dim numeroInstruct As Integer
    numeroInstruct = 2
    
    '/** ARRANGEMENTS ALAPHABETIQUE A 4 LETTRES **/

    For compterDonnees = 0 To 1
    
        'Créer une nouvelle feuille et lui donner un nom pour chaque lettre
        nomFeuille = alphabet(compteurFeuille)
        Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = nomFeuille
    
        For compterDonnees2 = 0 To 25

            For compterDonnees3 = 0 To 25
            
                For compterDonnees4 = 0 To 25
                    
                    'Intégrer les arrangements alphabétiques
                    'Worksheets
                    
                    
                    Sheets(nomFeuille).Range("" & lettreA & alphabet(colonne) & numeroCell & "").Value = alphabet(compterDonnees) + alphabet(compterDonnees2) + alphabet(compterDonnees3) + alphabet(compterDonnees4)
                    
                    'MsgBox "on est dans la boucle!"
                    'Intégrer les quantités associés
                    'Sheets("usagers").Activate
                    Sheets(nomFeuille).Range("" & lettreA & alphabet(colonne + 1) & numeroCell & "").Select
                    ActiveCell.FormulaLocal = "=NB.SI(usagers!A1:A210;""" & alphabet(compterDonnees) + alphabet(compterDonnees2) + alphabet(compterDonnees3) + alphabet(compterDonnees4) & "*"")"
                    
                    '/*******/
                    
                    'Compter noms et coloriser cellules
                    If compterNoms < divDonneesParInstruct Then
                        compterNoms = compterNoms + Sheets(nomFeuille).Range("" & lettreA & alphabet(colonne + 1) & numeroCell & "").Value
                        Sheets(nomFeuille).Range("" & lettreA & alphabet(colonne) & numeroCell & "").Interior.Color = RGB(colorCell, colorCell2, colorCell3)
                    Else
                        'MsgBox compterNoms
                        'Intégrer quantité de dossier dans colonne F (2 à 8)
                        Sheets(nomFeuille).Range("F" & numeroInstruct & "").Value = compterNoms
                        'Intégrer arrangement alphabétique dans colonne G (2 à 8)
                        Sheets(nomFeuille).Range("G" & numeroInstruct & "").Value = alphabet(compterDonnees) + alphabet(compterDonnees2) + alphabet(compterDonnees3) + alphabet(compterDonnees4 - 1)
                        'Passer à la cellule suivante pour le prochain tour
                        numeroInstruct = numeroInstruct + 1
                        
                        'Reset du comptage de noms
                        compterNoms = 0
                        compterNoms = compterNoms + Sheets(nomFeuille).Range("" & lettreA & alphabet(colonne + 1) & numeroCell & "").Value
                        colorCell = colorCell + 6
                        colorCell2 = colorCell2 + 30 * 1.15
                        colorCell3 = colorCell3 + 50 * 0.7
                        'colorCell = colorCell + colorCells
                        Sheets(nomFeuille).Range("" & lettreA & alphabet(colonne) & numeroCell & "").Interior.Color = RGB(colorCell, colorCell2, colorCell3)
                    End If
                    
                    '/*******/
                    
                    'On passe à la cellule suivante
                    numeroCell = numeroCell + 1
                    
                    'Si on dépasse la colonne Z
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
        
        'On passe à la feuille suivante
        compteurFeuille = compteurFeuille + 1
               
    Next
    'Si quantité n'est pas exact à la quantité de dossier théorique par instructrice (divDonneesParInstruct)
    'Stopper boucle et Ajouter un arrangement supplémentaire
    'Insérer les arrangements alpahbétiques limites pour chaque instructrice : de x lettre(s) à y lettre(s)

End Sub
