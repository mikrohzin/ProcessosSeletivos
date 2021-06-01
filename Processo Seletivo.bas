Attribute VB_Name = "Módulo1"
Sub main()
Dim Linhas As Integer, i As Long, Names() As String, Emails() As String, j As Integer, Dia As String
Dim objeto_outlook As Object
Dim Mail As Object
Set objeto_outlook = CreateObject("Outlook.Application")
Set Mail = objeto_outlook.createitem(0)

Call CLinha(Linhas)
Call Nomes(Names, Linhas)
Call Email(Emails, Linhas)
Call DiaNow(Dia)

With Mail
    Do While j < UBound(Emails)
        Mail.display
        Mail.to = Emails(j)
        Mail.Subject = "Retorno Processo Seletivo"
        Mail.Body = Dia & Names(j) & Chr(13) & Range("M2").Value
        Mail.Send
        j = j + 1
    Loop
End With
MsgBox ("Concluido")

End Sub
Function CLinha(Linhas)

Range("A2").Select
Linhas = Range(Selection, Selection.End(xlDown)).Rows.Count

End Function

Function DiaNow(Dia)

If Time > TimeValue("00:00:00") And Time < TimeValue("12:00:00") Then
    Dia = "Bom dia, "
ElseIf Time >= TimeValue("12:00:00") And Time < TimeValue("18:00:00") Then
    Dia = "Boa tarde, "
Else
    Dia = "Boa noite,"
End If

End Function

Function Email(Emails, Linhas)

ReDim Emails(Linhas)
Range("B2").Select

Do While i < Linhas
    If UCase(Selection.Offset(0, 1).Value) <> UCase("Aprovado") Then
        Emails(j) = Selection.Value
        Selection.Offset(1, 0).Select
        i = i + 1
        j = j + 1
    Else
        Selection.Offset(1, 0).Select
        i = i + 1
    End If
Loop
ReDim Preserve Emails(j)
End Function

Function Nomes(Names, Linhas)

ReDim Names(Linhas)
Range("A2").Select
Do While i < Linhas
    If UCase(Selection.Offset(0, 2).Value) <> UCase("Aprovado") Then
        Names(j) = Selection.Value
        Selection.Offset(1, 0).Select
        i = i + 1
        j = j + 1
    Else
        Selection.Offset(1, 0).Select
        i = i + 1
    End If
Loop
ReDim Preserve Names(j)
End Function
