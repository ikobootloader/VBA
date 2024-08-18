Option Explicit

' Structure pour stocker les informations du ticket
Type Ticket
    ID As Long
    Description As String
    Statut As String
    DateCreation As Date
    DateMiseAJour As Date
End Type

' Collection pour stocker tous les tickets
Dim Tickets As New Collection

' Fonction pour créer un nouveau ticket
Function CreerTicket(Description As String) As Long
    Dim NouveauTicket As Ticket
    
    NouveauTicket.ID = Tickets.Count + 1
    NouveauTicket.Description = Description
    NouveauTicket.Statut = "Ouvert"
    NouveauTicket.DateCreation = Now
    NouveauTicket.DateMiseAJour = Now
    
    Tickets.Add NouveauTicket
    
    CreerTicket = NouveauTicket.ID
End Function

' Fonction pour mettre à jour le statut d'un ticket
Sub MettreAJourStatut(ID As Long, NouveauStatut As String)
    Dim T As Ticket
    
    For Each T In Tickets
        If T.ID = ID Then
            T.Statut = NouveauStatut
            T.DateMiseAJour = Now
            Exit Sub
        End If
    Next T
    
    MsgBox "Ticket non trouvé.", vbExclamation
End Sub

' Fonction pour afficher tous les tickets
Sub AfficherTickets()
    Dim T As Ticket
    Dim Ws As Worksheet
    Dim Row As Long
    
    Set Ws = ThisWorkbook.Sheets("Tickets")
    Ws.Cells.Clear
    
    ' En-têtes
    Ws.Cells(1, 1) = "ID"
    Ws.Cells(1, 2) = "Description"
    Ws.Cells(1, 3) = "Statut"
    Ws.Cells(1, 4) = "Date de création"
    Ws.Cells(1, 5) = "Dernière mise à jour"
    
    Row = 2
    For Each T In Tickets
        Ws.Cells(Row, 1) = T.ID
        Ws.Cells(Row, 2) = T.Description
        Ws.Cells(Row, 3) = T.Statut
        Ws.Cells(Row, 4) = T.DateCreation
        Ws.Cells(Row, 5) = T.DateMiseAJour
        Row = Row + 1
    Next T
    
    Ws.Columns.AutoFit
End Sub
