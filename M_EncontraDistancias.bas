Attribute VB_Name = "M_EncontraDistancias"
Option Explicit

Function GetXML(url As String) As String

Dim XMLHttpRequest As Object 'xmlhttp

Set XMLHttpRequest = Interaction.CreateObject("MSXML2.xmlhttp") ' New MSXML2.xmlhttp
XMLHttpRequest.Open "GET", url, False
XMLHttpRequest.setRequestHeader "Content-Type", "text/xml"
XMLHttpRequest.send
GetXML = XMLHttpRequest.responseText

End Function

Sub BuscaIEdistUpdate()

Dim inicio As String, MyStr As String, swap1 As String, swap2 As String, strAddress As String
Dim linha As Integer, j As Integer, n1 As Long, tam As Long, n2 As Long
Dim kilometragem As Long, tempos As Long
Dim IE As InternetExplorer 'Declara objeto internet explorer
Dim isOK As Boolean

linha = 5

With Plan1
  
  .Activate

  While .Cells(linha, 1).Value <> ""
    
    'troca virgula por ponto
    swap1 = Strings.Replace(.Cells(linha, 1), " ", "+") 'origem
    swap2 = Strings.Replace(.Cells(linha, 2), " ", "+") 'destino
    
    swap1 = RetirarAcento(swap1)
    swap2 = RetirarAcento(swap2)
    
    '\\ Define a consulta no google maps
    strAddress = "http://maps.google.com/maps/api/directions/xml?origin=" & swap1 & ",Brasil&destination=" & swap2 & ",Brasil&sensor=false"
    
    '\\ lê o xml
    MyStr = GetXML(strAddress)
    
    'Verificador se o limite de pesquisas foi atingido
    ' Se der pau, tem que cadastrar outro codigo de acesso ao gmaps
    If Strings.InStr(1, MyStr, "OVER_QUERY_LIMIT") > 0 Then
      Set IE = Nothing
      Interaction.MsgBox ("Limite de pesquisa do google atingido")
      Exit Sub
    End If
    
    isOK = False
    
    'acha o ultimo duration
    While isOK = False
      n1 = Strings.InStr(1, MyStr, "<duration>")
      tam = Strings.Len(MyStr)
      MyStr = Strings.Mid(MyStr, n1 + 5, tam - n1)
      If n1 = 0 Then
        isOK = True
      End If
    Wend

    n1 = Strings.InStr(1, MyStr, "value")
    n2 = Strings.InStr(n1 + 5, MyStr, "/value")
    tempos = Strings.Mid(MyStr, n1 + 6, n2 - 1 - (n1 + 6))
                
    'acha ultimo distance
    isOK = False
    While isOK = False
      n1 = Strings.InStr(1, MyStr, "<distance>")
      tam = Strings.Len(MyStr)
      MyStr = Strings.Mid(MyStr, n1 + 5, tam - n1)
      If n1 = 0 Then
        isOK = True
      End If
    Wend
    
    'Encontra a palavra "text"
    n1 = Strings.InStr(1, MyStr, "value")
    n2 = Strings.InStr(n1 + 5, MyStr, "/value")
    kilometragem = Strings.Mid(MyStr, n1 + 6, n2 - 1 - (n1 + 6))
                
    '\\ grava os resultados
    .Cells(linha, 4) = kilometragem / 1000
    .Cells(linha, 5) = tempos / 60
    
    linha = linha + 1
    
  Wend

End With

End Sub

Function RetirarAcento(Palavra As String) As String
    
  RetirarAcento = Strings.Replace(Palavra, "ç", "c")
  RetirarAcento = Strings.Replace(RetirarAcento, "Ç", "C")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "á", "a")
  RetirarAcento = Strings.Replace(RetirarAcento, "à", "a")
  RetirarAcento = Strings.Replace(RetirarAcento, "ã", "a")
  RetirarAcento = Strings.Replace(RetirarAcento, "â", "a")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "é", "e")
  RetirarAcento = Strings.Replace(RetirarAcento, "ê", "e")
  RetirarAcento = Strings.Replace(RetirarAcento, "è", "e")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "í", "i")
  RetirarAcento = Strings.Replace(RetirarAcento, "ì", "i")
  RetirarAcento = Strings.Replace(RetirarAcento, "î", "i")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "ó", "o")
  RetirarAcento = Strings.Replace(RetirarAcento, "õ", "o")
  RetirarAcento = Strings.Replace(RetirarAcento, "ô", "o")
  RetirarAcento = Strings.Replace(RetirarAcento, "ò", "o")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "ú", "u")
  RetirarAcento = Strings.Replace(RetirarAcento, "ù", "u")
  RetirarAcento = Strings.Replace(RetirarAcento, "û", "u")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "Á", "A")
  RetirarAcento = Strings.Replace(RetirarAcento, "À", "A")
  RetirarAcento = Strings.Replace(RetirarAcento, "Ã", "A")
  RetirarAcento = Strings.Replace(RetirarAcento, "Â", "A")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "É", "E")
  RetirarAcento = Strings.Replace(RetirarAcento, "Ê", "E")
  RetirarAcento = Strings.Replace(RetirarAcento, "È", "E")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "Í", "I")
  RetirarAcento = Strings.Replace(RetirarAcento, "Ì", "I")
  RetirarAcento = Strings.Replace(RetirarAcento, "Î", "I")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "Ó", "O")
  RetirarAcento = Strings.Replace(RetirarAcento, "Õ", "O")
  RetirarAcento = Strings.Replace(RetirarAcento, "Ô", "O")
  RetirarAcento = Strings.Replace(RetirarAcento, "Ò", "O")
  
  RetirarAcento = Strings.Replace(RetirarAcento, "Ú", "U")
  RetirarAcento = Strings.Replace(RetirarAcento, "Ù", "U")
  RetirarAcento = Strings.Replace(RetirarAcento, "Û", "U")
       
End Function

Sub BuscaIEdist(linha As Integer)

Dim inicio As String, MyStr As String, distTemp As String
Dim swap1 As String, swap2 As String, swap3 As String, swap4 As String
Dim n1 As Long, n2 As Long, tam As Long
Dim IE As Object

'Cria o objeto para navegação Internet Explorer
Set IE = Interaction.CreateObject("InternetExplorer.Application")

'troca espaço por positivo
swap1 = Replace(Cells(linha, 1), " ", "+")
swap2 = Replace(Cells(linha, 2), " ", "+")

'retira acentos
swap1 = RetirarAcento(swap1)
swap2 = RetirarAcento(swap2)

'cria o objeto para visualização do maps
Set IE = CreateObject("InternetExplorer.Application")
IE.Navigate "http://maps.google.com/maps?f=d&saddr=" & swap1 & "&daddr=" & swap2

'aguarda até que a pagina seja totalmente carregada
Do Until Not IE.Busy And IE.ReadyState = READYSTATE_COMPLETE
Loop

IE.Visible = True
        
End Sub
