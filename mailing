' GPT4o

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    Dim Mail_Subject As String
    Dim Mail_Body As String
    Dim Mail_Object As String
    Dim Mail_Objectives As String
    Dim Mail_Participants As String
    Dim Recipients As String
    
    ' Définir la colonne à surveiller, par exemple, la colonne F pour "Date de réalisation"
    If Not Intersect(Target, Me.Columns("F")) Is Nothing Then
        For Each cell In Target
            ' Si une cellule dans la colonne surveillée est modifiée
            If cell.Value <> "" Then
                ' Extraire les données nécessaires
                Mail_Object = Me.Cells(cell.Row, "B").Value ' Objet
                Mail_Objectives = Me.Cells(cell.Row, "C").Value ' Objectifs
                Mail_Participants = Me.Cells(cell.Row, "D").Value ' Participants
                
                ' Obtenir les adresses e-mail des participants
                Recipients = GetEmailAddresses(Mail_Participants)
                
                ' Demander à l'utilisateur s'il veut envoyer un mail
                Dim User_Response As VbMsgBoxResult
                User_Response = MsgBox("Voulez-vous prévenir les participants ?", vbYesNo + vbQuestion, "Confirmation d'envoi")
                
                ' Si l'utilisateur choisit "Oui"
                If User_Response = vbYes Then
                    ' Créer le sujet et le corps du mail
                    Mail_Subject = "Notification de Projet: " & Mail_Object
                    Mail_Body = "Bonjour," & vbNewLine & vbNewLine & _
                                "L'objectif suivant a été mis à jour :" & vbNewLine & _
                                Mail_Objectives & vbNewLine & vbNewLine & _
                                "Cordialement," & vbNewLine & _
                                "Votre équipe de projet"
                    
                    ' Envoyer le mail
                    SendMail Recipients, Mail_Subject, Mail_Body
                End If
            End If
        Next cell
    End If
End Sub

Function GetEmailAddresses(Participants As String) As String
    Dim EmailDict As Object
    Set EmailDict = CreateObject("Scripting.Dictionary")
    
    ' Ajouter les combinaisons de participants et leurs adresses e-mail respectives
    EmailDict.Add "marc+sophie", "marc@example.com;sophie@example.com"
    EmailDict.Add "alice+bob", "alice@example.com;bob@example.com"
    EmailDict.Add "marc+sophie+alice+bob", "marc@example.com;sophie@example.com;alice@example.com;bob@example.com"
    ' Ajouter d'autres combinaisons si nécessaire

    ' Retourner les adresses e-mail correspondantes
    If EmailDict.Exists(Participants) Then
        GetEmailAddresses = EmailDict(Participants)
    Else
        GetEmailAddresses = ""
    End If
    
    ' Libérer l'objet
    Set EmailDict = Nothing
End Function

Sub SendMail(Recipients As String, Subject As String, Body As String)
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    
    ' Créer une instance d'Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    With OutlookMail
        .To = Recipients ' Adresses mail des participants
        .CC = "" ' Adresse mail en copie
        .BCC = "" ' Adresse mail en copie cachée
        .Subject = Subject
        .Body = Body
        .Display ' Afficher le mail avant envoi
        '.Send ' Envoyer directement le mail (décommenter pour activer)
    End With
    
    ' Libérer les objets
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub
