Attribute VB_Name = "ARENA_CONFIGSALAS"
Sub toogleConfigSalas()
    If Sheets("CONFIG-SALAS").Visible = True Then
        Sheets("CONFIG-SALAS").Visible = False
        Sheets("CONFIG-QTD").Visible = False
    Else
        Sheets("CONFIG-SALAS").Visible = True
        Sheets("CONFIG-QTD").Visible = True
    End If
End Sub
