VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub macroRelatorio()

Dim linha As Integer
Dim ultimaLinha As Integer
Dim ws As Worksheet

linha = 1
ultimaLinha = Range("A1").End(xlDown).Row

Range("F2:F" & ultimaLinha).Copy
Range("J1").PasteSpecial

' tirar a "cobrinha" do copiar
Application.CutCopyMode = False

' remover duplicadas
Set ws = Worksheets("Planilha1")
ws.Range("$J$1:$J$" & ultimaLinha).RemoveDuplicates Columns:=1, Header:=xlNo


ultimaLinha = Range("J1").End(xlDown).Row


For i = 1 To ultimaLinha

    ' vai adicionar uma nova aba (sheet)
    Sheets.Add after:=ActiveSheet
    
    ' colocar nome na nova planilha
    ' primeiro argumento de Cells é a linha, e segundo a coluna
    ActiveSheet.Name = ws.Cells(i, 10)
    
    ' copiando o cabeçalho
    Sheets("Planilha1").Range("A1:F1").Copy
    
    ' colar cabeçalho na nova aba (sheet)
    ActiveSheet.Range("A1").PasteSpecial
    

Next i
    
' limpar dados colados anteriormente
ws.Range("J:J").Clear

linha = 2

'ultimaLinha = Range("A1").End(xlDown).Row

'While ws.Cells(linha, 1) <> ""
For i = 2 To 200

    ' copiar todos os dados da tabela
    ws.Range("A" & linha & ":F" & linha).Copy
    
    ' pegando estado e selecionando planilha
    planilhaEstado = ws.Cells(linha, 6)
    'Sheets(planilhaEstado).Select
    
    'copiar na ultima linha da planilha selecionada
    Sheets(planilhaEstado).Range("A10000").End(xlUp).Offset(1, 0).PasteSpecial
    
    ws.Select
    
    linha = linha + 1
    
Next i
'Wend

' navegar por cada aba (sheets) e fazer auto fit
For Each aba In ThisWorkbook.Sheets
    aba.Columns("A:F").AutoFit
Next

MsgBox "Relatórios criados!", vbInformation, "Sucesso"

End Sub
