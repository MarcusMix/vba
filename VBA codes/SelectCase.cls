VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub selectCase()

    Dim numero As Integer
    
    numero = InputBox("Digite um n�mero inteiro no intervalo de 1 a 10")
    
    Select Case numero
    
        Case 1 To 5
        
            MsgBox "N�mero entre 1 e 5"
            
        Case 6 To 8
        
            MsgBox "N�mero entre 6 e 8"
            
        Case 9 To 10
        
            MsgBox "N�mero entre 9 e 10"
            
        Case Else
        
            MsgBox "N�mero n�o encontrado"
            
    End Select

End Sub
