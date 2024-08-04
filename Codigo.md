CRIAÇÃO DA FICHA TÉCNICA DOS PRODUTOS DA SINTO DO BRASIL. 

Objetivo: Pela necessidade da empresa solicitou em mudar o layout da ficha técnica, propôs substituir a ferramenta Word pelo Excel 
para deixar o processo automático e fácil.

1. Foi criado o layout da ficha técnica:

![image](https://github.com/user-attachments/assets/5b763c31-31ef-4f04-b587-fc9ca11a9309)

2. Em uma segunda planilha foi criado todo o banco de dados com todos os dados e especeficações dos materiais.

3. Por meio do Visual Basic, foi criado o layout do formulario de "Emissão de ficha Técnica".

![image](https://github.com/user-attachments/assets/399c02e6-ebcd-4644-845f-b88cecd28d9c)

4. Criado o layout do formulário "Cadastrar novos produtos"

![image](https://github.com/user-attachments/assets/aa05d216-21fd-4edf-985b-f8862ddc6bce)

5. Criação dos códigos para o formulario "Emissão de ficha técnica"

•	Códigos para ser carregado ao iniciar o formulario. Utilizado para adicionar os dados do idioma e dos 
materiais da Sinto dentro do ComboBox.

```

Private Sub UserForm_Initialize()

Dim Lin As Integer

Lin = 3

    Do Until PlanBancoDeDados.Cells(Lin, 58) = ""
    
         CmbIdioma.AddItem PlanBancoDeDados.Cells(Lin, 58)
         
         Lin = Lin + 1
     
     Loop

Lin = 3
     
     Do Until PlanBancoDeDados.Cells(Lin, 62) = ""
     
     CmbMaterial.AddItem PlanBancoDeDados.Cells(Lin, 62)
     
        Lin = Lin + 1
     
     Loop

End Sub
```
•	Códigos que serão executados ao clicar no botão "Gerar ficha técnica"

```
Private Sub cmdEmissão_Click()

    'Criar uma variavel como string
    Dim strFilename As String
    
    
    'Se o campo do idioma estiver vazio, mandar uma mensagem para o usuário digitar
    If CmbIdioma = "" Then
    MsgBox "Selecione o idioma desejado"
    Exit Sub

    'Se o campo do idioma estiver diferente de português, inglês e espanhol mostrar uma mensagem de valor inválido
    ElseIf CmbIdioma.ListIndex = -1 Then
    MsgBox "valor inválido, selecione uma lingua válida"
    CmbIdioma.SetFocus
    CmbIdioma = ""
    Exit Sub
    End If
    
    'Se o campo do material estiver vazio, mandar uma mensagem para o usuário digitar
    If CmbMaterial = "" Then
    MsgBox "Selecione o material desejado"
    Exit Sub
    
    'Quando o usuário não seleciona nenhum item, ou então ele escreve um texto que não está entre os itens
    'da lista, o índice que o combobox retorna é um negativo. (combobox.listIndex = -1)
    ElseIf CmbMaterial.ListIndex = -1 Then
    MsgBox "Selecione um material existente"
    CmbMaterial.SetFocus
    CmbMaterial = ""
    Exit Sub
    End If
    
    'Define  variáveis
    Dim Tipo_material As String
    Dim Material_QtdMalhas As String
    
    'Define o valor da variável como sendo o valor procurado na vertical da matriz do banco de dados com a coluna 3
    Material_QtdMalhas = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("AK9:AN655"), 3, False)
    'Define o valor da variável como sendo o valor procurado na vertical da matriz do banco de dados com a coluna 2
     Tipo_material = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("AK9:AN655"), 2, False)
    
    Application.ScreenUpdating = False      'Desativa a atualização de página quando a macro está sendo executada
    Application.DisplayAlerts = False       'Desativa mensagens de alerta e aviso no Excel
    PlanModelo.Select                       'Seleciona a planilha Modelo
    Range("M24:BS26").ClearContents         'Limpa fórmulas e valores do intervalo
    Range("B24:BS26").UnMerge               'Desmeclar a coleção de células
    Range("B24:BS26").Select                'Seleciona o intervalo
    
    'Adicionar o material na ficha técnica
    Range("AA10:BS10") = CmbMaterial
    'Adicionar o idioma "ficha tecnica"
    Range("B7:BS7") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 2, False)
    'Adicionar o idioma "material"
     Range("B10:G10") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 3, False)
    'Adicionar o idioma "granalha de aço alto carbono"
     Range("H10:Z10") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 4, False)
    'Adicionar especificação da composição química
    Range("B13:BS13") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 2, False)
    'Adicionar especificação da dureza
    Range("B19:BS19") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 3, False)
    'Adicionar especificação da granulometria
    Range("B23:BS23") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 4, False)
    'Adicionar especificação da densidade
    Range("B29:BS29") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 5, False)
    'Adicionar especificação da macroestrutura
    Range("B33:BS33") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 6, False)
    'Adicionar especificação da microestrutura
    Range("B38:BS38") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 7, False)
    'Adicionar especificação do não magnético
    Range("B42:BS42") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 8, False)
    'Adicionar especificação da norma SAE
    Range("B50:BS50") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 9, False)
    'Adicionar o idioma "carbono"
     Range("M14:X14") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 5, False)
    'Adicionar o idioma "Manganês"
     Range("Y14:AJ14") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 6, False)
    'Adicionar o idioma "Silício"
     Range("AK14:AU14") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 7, False)
    'Adicionar o idioma "Fósforo"
     Range("AV14:BG14") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 8, False)
    'Adicionar o idioma "Enxofre"
     Range("BH14:BS14") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 9, False)
    'Adicionar o idioma "Especificação (%) composição química"
     Range("B16:L16") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 10, False)
    'Adicionar o idioma "Especificação dureza"
     Range("B20:L20") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 11, False)
    'Adicionar o idioma "Malha"
     Range("B24:L24") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 12, False)
    'Adicionar o idioma "Abertura"
     Range("B25:L25") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 13, False)
    'Adicionar o idioma "Especificação granulometria"
     Range("B26:L26") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 14, False)
    'Adicionar o idioma "Especificação (g/cm³) densidade"
     Range("B30:O30") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 16, False)
    'Adicionar o idioma "especificação macroestrutura"
     Range("B35:L35") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 22, False)
     'Adicionar o idioma "especificação microestrutura"
     Range("B39:L39") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 28, False)
    'Adicionar o idioma "especificação não magnético"
     Range("B43:l43") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 29, False)
    'Adicionar o idioma "Material não magnético máximo 1%"
     Range("M43:BS43") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 30, False)
    'Adicionar o idioma "Especificação técnica em conformidade com:"
     Range("B46:BS46") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 31, False)
    'Adicionar especificação "90% das partículas devem estar entre xx a xx HRC"
    Range("M20:BS20") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 10, False)
    'Adicionar especificação "Mín. x.xx (densidade)"
    Range("P30:BS30") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 32, False)
    'Adicionar especificação "Mín. x.xx (densidade)"
    Range("M39:BS39") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 34, False)
    
    With Selection.Borders(xlEdgeLeft)      'Seleciona a borda da esquerda para altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    With Selection.Borders(xlEdgeTop)       'Seleciona a borda de cima para altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    With Selection.Borders(xlEdgeBottom)    'Seleciona a borda de baixo para altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    With Selection.Borders(xlEdgeRight)     'Seleciona a borda da direita para altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    With Selection.Borders(xlInsideVertical) 'Seleciona as bordas verticais internas e altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    With Selection.Borders(xlInsideHorizontal) 'Seleciona as bordas horizontais internas e altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    End With:  End With:  End With:   End With:     End With:      End With
    
    Range("M34:BS35").ClearContents         'Limpa fórmulas e valores do intervalo
    Range("B34:BS35").UnMerge               'Desmeclar a coleção de células
    Range("B34:BS35").Select                'Seleciona o intervalo
    With Selection.Borders(xlEdgeLeft)      'Seleciona a borda da esquerda para altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    With Selection.Borders(xlEdgeTop)       'Seleciona a borda de cima para altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    With Selection.Borders(xlEdgeBottom)    'Seleciona a borda de baixo para altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    With Selection.Borders(xlEdgeRight)     'Seleciona a borda da direita para altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    With Selection.Borders(xlInsideVertical) 'Seleciona as bordas verticais internas e altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    With Selection.Borders(xlInsideHorizontal) 'Seleciona as bordas horizontais internas e altera suas propriedades
        .LineStyle = xlContinuous           'Estilo da linha continua
        .Weight = xlThin                    'espessura = fino
    End With:  End With:  End With:   End With:     End With:      End With
    
    'Se a variável encontrar no valor procurado "Shot 2 Malhas" fará o código abaixo
    If Material_QtdMalhas = "2 Malhas" Then
    Range("B24:L24").Merge          'Mescla a seleção
    Range("B25:L25").Merge          'Mescla a seleção
    Range("B26:L26").Merge          'Mescla a seleção
    Range("M24:AO24").Merge         'Mescla a seleção
    Range("M25:AO25").Merge         'Mescla a seleção
    Range("M26:AO26").Merge         'Mescla a seleção
    Range("AP24:BS24").Merge        'Mescla a seleção
    Range("AP25:BS25").Merge        'Mescla a seleção
    Range("AP26:BS26").Merge        'Mescla a seleção
        
    'Adicionada o valor da primeira malha
    Range("M24:AO24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 11, False)
    'Adicionada o valor da primeira abertura
     Range("M25:AO25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 12, False)
    'Adicionada o valor da primeira especificação
    Range("M26:AO26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 13, False)
    'Adicionada o valor da segunda malha
    Range("AP24:BS24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 14, False)
    'Adicionada o valor da segunda abertura
    Range("AP25:BS25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 15, False)
    'Adicionada o valor da segunda especificação
    Range("AP26:BS26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 16, False)
    
    'Se a variável encontrar no valor procurado "3 Malhas" fará o código abaixo
    ElseIf Material_QtdMalhas = "3 Malhas" Then
    Range("B24:L24").Merge          'Mescla a seleção
    Range("B25:L25").Merge          'Mescla a seleção
    Range("B26:L26").Merge          'Mescla a seleção
    Range("M24:AE24").Merge         'Mescla a seleção
    Range("AF24:AY24").Merge        'Mescla a seleção
    Range("AZ24:BS24").Merge        'Mescla a seleção
    Range("M25:AE25").Merge         'Mescla a seleção
    Range("AF25:AY25").Merge        'Mescla a seleção
    Range("AZ25:BS25").Merge        'Mescla a seleção
    Range("M26:AE26").Merge         'Mescla a seleção
    Range("AF26:AY26").Merge        'Mescla a seleção
    Range("AZ26:BS26").Merge        'Mescla a seleção
                                     
    'Adicionada o valor da primeira malha
    Range("M24:AE24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 11, False)
    'Adicionada o valor da primeira abertura
     Range("M25:AE25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 12, False)
    'Adicionada o valor da primeira especificação
    Range("M26:AE26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 13, False)
    'Adicionada o valor da segunda malha
    Range("AF24:AY24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 14, False)
    'Adicionada o valor da segunda abertura
    Range("AF25:AY25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 15, False)
    'Adicionada o valor da segunda especificação
    Range("AF26:AY26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 16, False)
    'Adicionada o valor da terceira malha
    Range("AZ24:BS24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 17, False)
    'Adicionada o valor da terceira abertura
    Range("AZ25:BS25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 18, False)
    'Adicionada o valor da terceira especificação
    Range("AZ26:BS26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 19, False)
    
    'Se a variável encontrar no valor procurado "4 Malhas" fará o código abaixo
    ElseIf Material_QtdMalhas = "4 Malhas" Then
    Range("B24:L24").Merge          'Mescla a seleção
    Range("B25:L25").Merge          'Mescla a seleção
    Range("B26:L26").Merge          'Mescla a seleção
    Range("M24:Z24").Merge          'Mescla a seleção
    Range("M25:Z25").Merge          'Mescla a seleção
    Range("M26:Z26").Merge          'Mescla a seleção
    Range("AA24:AO24").Merge        'Mescla a seleção
    Range("AA25:AO25").Merge        'Mescla a seleção
    Range("AA26:AO26").Merge        'Mescla a seleção
    Range("AP24:BD24").Merge        'Mescla a seleção
    Range("AP25:BD25").Merge        'Mescla a seleção
    Range("AP26:BD26").Merge        'Mescla a seleção
    Range("BE24:BS24").Merge        'Mescla a seleção
    Range("BE25:BS25").Merge        'Mescla a seleção
    Range("BE26:BS26").Merge        'Mescla a seleção
    
    'Adicionada o valor da primeira malha
    Range("M24:AE24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 11, False)
    'Adicionada o valor da primeira abertura
     Range("M25:AE25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 12, False)
    'Adicionada o valor da primeira especificação
    Range("M26:AE26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 13, False)
    'Adicionada o valor da segunda malha
    Range("AA24:AO24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 14, False)
    'Adicionada o valor da segunda abertura
    Range("AA25:AO25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 15, False)
    'Adicionada o valor da segunda especificação
    Range("AA26:AO26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 16, False)
    'Adicionada o valor da terceira malha
    Range("AP24:BD24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 17, False)
    'Adicionada o valor da terceira abertura
    Range("AP25:BD25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 18, False)
    'Adicionada o valor da terceira especificação
    Range("AP26:BD26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 19, False)
    'Adicionada o valor da quarta malha
    Range("BE24:BS24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 20, False)
    'Adicionada o valor da quarta abertura
    Range("BE25:BS25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 21, False)
    'Adicionada o valor da quarta especificação
    Range("BE26:BS26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 22, False)
    
    'Se a variável encontrar no valor procurado "5 Malhas" fará o código abaixo
    ElseIf Material_QtdMalhas = "5 Malhas" Then
    Range("B24:L24").Merge          'Mescla a seleção
    Range("B25:L25").Merge          'Mescla a seleção
    Range("B26:L26").Merge          'Mescla a seleção
    Range("M24:W24").Merge          'Mescla a seleção
    Range("M25:W25").Merge          'Mescla a seleção
    Range("M26:W26").Merge          'Mescla a seleção
    Range("X24:AI24").Merge         'Mescla a seleção
    Range("X25:AI25").Merge         'Mescla a seleção
    Range("X26:AI26").Merge         'Mescla a seleção
    Range("AJ24:AU24").Merge        'Mescla a seleção
    Range("AJ25:AU25").Merge        'Mescla a seleção
    Range("AJ26:AU26").Merge        'Mescla a seleção
    Range("AV24:BG24").Merge        'Mescla a seleção
    Range("AV25:BG25").Merge        'Mescla a seleção
    Range("AV26:BG26").Merge        'Mescla a seleção
    Range("BH24:BS24").Merge        'Mescla a seleção
    Range("BH25:BS25").Merge        'Mescla a seleção
    Range("BH26:BS26").Merge        'Mescla a seleção
        
    'Adicionada o valor da primeira malha
    Range("M24:W24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 11, False)
    'Adicionada o valor da primeira abertura
     Range("M25:W25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 12, False)
    'Adicionada o valor da primeira especificação
    Range("M26:W26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 13, False)
    'Adicionada o valor da segunda malha
    Range("X24:AI24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 14, False)
    'Adicionada o valor da segunda abertura
    Range("X25:AI25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 15, False)
    'Adicionada o valor da segunda especificação
    Range("X26:AI26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 16, False)
    'Adicionada o valor da terceira malha
    Range("AJ24:AU24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 17, False)
    'Adicionada o valor da terceira abertura
    Range("AJ25:AU25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 18, False)
    'Adicionada o valor da terceira especificação
    Range("AJ26:AU26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 19, False)
    'Adicionada o valor da quarta malha
    Range("AV24:BG24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 20, False)
    'Adicionada o valor da quarta abertura
    Range("AV25:BG25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 21, False)
    'Adicionada o valor da quarta especificação
    Range("AV26:BG26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 22, False)
    'Adicionada o valor da quinta malha
    Range("BH24:BS24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 23, False)
    'Adicionada o valor da quinta abertura
    Range("BH25:BS25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 24, False)
    'Adicionada o valor da quinta especificação
    Range("BH26:BS26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 25, False)
    
    'Se a variável encontrar no valor procurado "6 Malhas" fará o código abaixo
    ElseIf Material_QtdMalhas = "6 Malhas" Then
    Range("B24:L24").Merge                  'mescla a seleção
    Range("B25:L25").Merge                  'mescla a seleção
    Range("B26:L26").Merge                  'mescla a seleção
    Range("M24:U24").Merge                  'mescla a seleção
    Range("M25:U25").Merge                  'mescla a seleção
    Range("M26:U26").Merge                  'mescla a seleção
    Range("V24:AE24").Merge                 'mescla a seleção
    Range("V25:AE25").Merge                 'mescla a seleção
    Range("V26:AE26").Merge                 'mescla a seleção
    Range("AF24:AO24").Merge                'mescla a seleção
    Range("AF25:AO25").Merge                'mescla a seleção
    Range("AF26:AO26").Merge                'mescla a seleção
    Range("AP24:AY24").Merge                'mescla a seleção
    Range("AP25:AY25").Merge                'mescla a seleção
    Range("AP26:AY26").Merge                'mescla a seleção
    Range("AZ24:BI24").Merge                'mescla a seleção
    Range("AZ25:BI25").Merge                'mescla a seleção
    Range("AZ26:BI26").Merge                'mescla a seleção
    Range("BJ24:BS24").Merge                'mescla a seleção
    Range("BJ25:BS25").Merge                'mescla a seleção
    Range("BJ26:BS26").Merge                'mescla a seleção

    'Adicionada o valor da primeira malha
    Range("M24:U24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 11, False)
    'Adicionada o valor da primeira abertura
     Range("M25:U25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 12, False)
    'Adicionada o valor da primeira especificação
    Range("M26:U26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 13, False)
    'Adicionada o valor da segunda malha
    Range("V24:AE24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 14, False)
    'Adicionada o valor da segunda abertura
    Range("V25:AE25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 15, False)
    'Adicionada o valor da segunda especificação
    Range("V26:AE26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 16, False)
    'Adicionada o valor da terceira malha
    Range("AF24:AO24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 17, False)
    'Adicionada o valor da terceira abertura
    Range("AF25:AO25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 18, False)
    'Adicionada o valor da terceira especificação
    Range("AF26:AO26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 19, False)
    'Adicionada o valor da quarta malha
    Range("AP24:AY24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 20, False)
    'Adicionada o valor da quarta abertura
    Range("AP25:AY25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 21, False)
    'Adicionada o valor da quarta especificação
    Range("AP26:AY26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 22, False)
    'Adicionada o valor da quinta malha
    Range("AZ24:BI24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 23, False)
    'Adicionada o valor da quinta abertura
    Range("AZ25:BI25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 24, False)
    'Adicionada o valor da quinta especificação
    Range("AZ26:BI26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 25, False)
    'Adicionada o valor da sexta malha
    Range("BJ24:BS24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 26, False)
    'Adicionada o valor da sexta abertura
    Range("BJ25:BS25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 27, False)
    'Adicionada o valor da sexta especificação
    Range("BJ26:BS26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 28, False)
    
    'Se a variável encontrar no valor procurado "7 Malhas" fará o código abaixo
    ElseIf Material_QtdMalhas = "7 Malhas" Then
    Range("B24:L24").Merge          'Mescla a seleção
    Range("B25:L25").Merge          'Mescla a seleção
    Range("B26:L26").Merge          'Mescla a seleção
    Range("M24:T24").Merge          'Mescla a seleção
    Range("M25:T25").Merge          'Mescla a seleção
    Range("M26:T26").Merge          'Mescla a seleção
    Range("U24:AB24").Merge         'Mescla a seleção
    Range("U25:AB25").Merge         'Mescla a seleção
    Range("U26:AB26").Merge         'Mescla a seleção
    Range("AC24:AJ24").Merge        'Mescla a seleção
    Range("AC25:AJ25").Merge        'Mescla a seleção
    Range("AC26:AJ26").Merge        'Mescla a seleção
    Range("AK24:AR24").Merge        'Mescla a seleção
    Range("AK25:AR25").Merge        'Mescla a seleção
    Range("AK26:AR26").Merge        'Mescla a seleção
    Range("AS24:BA24").Merge        'Mescla a seleção
    Range("AS25:BA25").Merge        'Mescla a seleção
    Range("AS26:BA26").Merge        'Mescla a seleção
    Range("BB24:BJ24").Merge        'Mescla a seleção
    Range("BB25:BJ25").Merge        'Mescla a seleção
    Range("BB26:BJ26").Merge        'Mescla a seleção
    Range("BK24:BS24").Merge        'Mescla a seleção
    Range("BK25:BS25").Merge        'Mescla a seleção
    Range("BK26:BS26").Merge        'Mescla a seleção
        
    'Adicionada o valor da primeira malha
    Range("M24:T24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 11, False)
    'Adicionada o valor da primeira abertura
     Range("M25:T25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 12, False)
    'Adicionada o valor da primeira especificação
    Range("M26:T26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 13, False)
    'Adicionada o valor da segunda malha
    Range("U24:AB24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 14, False)
    'Adicionada o valor da segunda abertura
    Range("U25:AB25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 15, False)
    'Adicionada o valor da segunda especificação
    Range("U26:AB26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 16, False)
    'Adicionada o valor da terceira malha
    Range("AC24:AJ24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 17, False)
    'Adicionada o valor da terceira abertura
    Range("AC25:AJ25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 18, False)
    'Adicionada o valor da terceira especificação
    Range("AC26:AJ26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 19, False)
    'Adicionada o valor da quarta malha
    Range("AK24:AR24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 20, False)
    'Adicionada o valor da quarta abertura
    Range("AK25:AR25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 21, False)
    'Adicionada o valor da quarta especificação
    Range("AK26:AR26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 22, False)
    'Adicionada o valor da quinta malha
    Range("AS24:BA24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 23, False)
    'Adicionada o valor da quinta abertura
    Range("AS25:BA25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 24, False)
    'Adicionada o valor da quinta especificação
    Range("AS26:BA26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 25, False)
    'Adicionada o valor da sexta malha
    Range("BB24:BJ24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 26, False)
    'Adicionada o valor da sexta abertura
    Range("BB25:BJ25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 27, False)
    'Adicionada o valor da sexta especificação
    Range("BB26:BJ26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 28, False)
    'Adicionada o valor da sétima malha
    Range("BK24:BS24") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 29, False)
    'Adicionada o valor da sétima abertura
    Range("BK25:BS25") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 30, False)
    'Adicionada o valor da sétima especificação
    Range("BK26:BS26") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 31, False)
    
    End If
    
    'Se a variável encontrar no valor procurado "Shot" fará o comando Shot
    If Tipo_material = "Shot" Then
    Range("B34:L34").Merge               'Mescla a seleção
    Range("B35:L35").Merge               'Mescla a seleção
    Range("M34:Z34").Merge               'Mescla a seleção
    Range("M35:Z35").Merge               'Mescla a seleção
    Range("AA34:AO34").Merge             'Mescla a seleção
    Range("AA35:AO35").Merge             'Mescla a seleção
    Range("AP34:BD34").Merge             'Mescla a seleção
    Range("AP35:BD35").Merge             'Mescla a seleção
    Range("BE34:BS34").Merge             'Mescla a seleção
    Range("BE35:BS35").Merge             'Mescla a seleção
       
    'Adicionada a primeira macroestrutura
    Range("M34:Z34") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 17, False)
    'Adicionada a primeira especificação
    Range("M35:Z35") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 24, False)
    'Adicionada a segunda macroestrutura
    Range("AA34:AO34") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 18, False)
    'Adicionada a segunda especificação
    Range("AA35:AO35") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 25, False)
    'Adicionada a terceira macroestrutura
    Range("AP34:BD34") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 19, False)
    'Adicionada a terceira especificação
    Range("AP35:BD35") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 26, False)
    'Adicionada a quarta macrestrutura
    Range("BE34:BS34") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 21, False)
    'Adicionada a quarta especificação
    Range("BE35:BS35") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 33, False)
    
    'Se a variável encontrar no valor procurado "Grit" fará:
    ElseIf Tipo_material = "Grit" Then
    
    Range("B34:L34").Merge               'Mescla a seleção
    Range("B35:L35").Merge               'Mescla a seleção
    Range("M34:AF34").Merge              'Mescla a seleção
    Range("M35:AF35").Merge              'Mescla a seleção
    Range("AG34:AZ34").Merge             'Mescla a seleção
    Range("AG35:AZ35").Merge             'Mescla a seleção
    Range("BA34:BS34").Merge             'Mescla a seleção
    Range("BA35:BS35").Merge             'Mescla a seleção

    'Adicionada a primeira macroestrutura
    Range("M34:AF34") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 19, False)
    'Adicionada a primeira especificação
    Range("M35:AF35") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 23, False)
    'Adicionada a segunda macroestrutura
    Range("AG34:AZ34") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 18, False)
    'Adicionada a segunda especificação
    Range("AG35:AZ35") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 24, False)
    'Adicionada a terceira macroestrutura
    Range("BA34:BS34") = WorksheetFunction.HLookup _
    (CmbIdioma, PlanBancoDeDados.Range("AP2:AR81"), 20, False)
    'Adicionada a terceira especificação
    Range("BA35:BS35") = WorksheetFunction.VLookup _
    (CmbMaterial & " " & CmbIdioma, PlanBancoDeDados.Range("C9:AJ655"), 33, False)
    
    End If
    
    'Variavel concatena o nome do material com o idioma
    strFilename = CmbMaterial & " " & CmbIdioma
    'salvar em PDF a ficha técnica
     ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    strFilename & ".pdf" _
    , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
    :=False, OpenAfterPublish:=True
    
    'Seleciona a planilha 1
    Plan1.Select
    
    Application.ScreenUpdating = True       'Reativa a atualização de página
    Application.DisplayAlerts = True        'Reativa mensagens de alerta e aviso no Excel
    
    'Mostra a mensagem para o usario
    MsgBox "Ficha tecnica emitida com sucesso", vbExclamation, "Ficha Técnica"
    
    'Fecha o formulário.
    Unload Me
    
    'Encerrar o código
   End Sub
```

6. Criado um botão dentro do formulário 'Emissão de ficha técnica' para realizar o cadastro de novos materiais

![image](https://github.com/user-attachments/assets/66135535-8060-4583-a316-dc4da689cada)

• Ao clicar nesse botão, será aberto uma requisição de senha para acessar o formulario de cadastro. 

![image](https://github.com/user-attachments/assets/714abc74-be86-406d-b630-a7db2830c5d4)

•	Códigos que serão executados ao clicar no botão

```
Private Sub CmdCadastro_Click()

'Define a senha como uma string

Dim senha As String
    
    'Usário digitar a senha de acesso"
    senha = InputBox("Digite a senha para ter acesso a esse formulário", "Senha requerida")
    
    'Se o usuario acertar a senha abrirá o fórmulario de cadastro
    If senha = "XXXXXX" Then
        fmfCadastrar.Show
       
    'Se o usuário errar a senha, aparecerá uma mensagem de senha inválida
    Else
        MsgBox "Senha incorreta", vbCritical, "Acesso ao cadastro"
    End If
    
'Encerrar o código
End Sub
```
7. Ao acertar a senha, será exibido o formulario de cadastro de novos materiais. Abaixo os códigos de como funciona tal formulario.

•	Códigos que serão executados ao iniciar o formulário

```
Private Sub UserForm_Initialize()
        
    'Adicionar dados para o campo de composição química
    CmbCompQuím.AddItem "Composição Química (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)"
    CmbCompQuím.AddItem "Composição Química (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)"
         
    'Adicionar dados para o campo de dureza
    CmbDureza.AddItem "Dureza (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)"
    CmbDureza.AddItem "Dureza (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)"
    CmbDureza.AddItem "Dureza (Baseado na norma BRS)"
    
    'Adicionar dados para o campo de granulometria
    CmbGranulometria.AddItem "Granulometria (Baseado na norma J444  - Cast Shot and Grit Size Specifications for Peening and Cleaning)"
    CmbGranulometria.AddItem "Granulometria (Baseado na norma BRS)"
         
    'Adicionar dados para o campo de composição densidade
    CmbDensidade.AddItem "Densidade (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)"
    CmbDensidade.AddItem "Densidade (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)"
    CmbDensidade.AddItem "Densidade (Baseado na norma BRS)"

    'Adicionar dados para o campo de macroestrutura
    CmbMacroestrutura.AddItem "Macroestrutura (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)"
    CmbMacroestrutura.AddItem "Macroestrutura (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)"
        
    'Adicionar dados para o campo de microestrutura"
    CmbMicroestrutura.AddItem "Microestrutura (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)"
    CmbMicroestrutura.AddItem "Microestrutura (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)"
         
    'Adicionar dados para o campo de não magnéticos
    CmbNãoMagnético.AddItem "Não magnéticos (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)"
    CmbNãoMagnético.AddItem "Não magnéticos (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)"
         
    'Adicionar dados para o campo SAE
    CmbSAE.AddItem " - SAE J827 - High-Carbon Cast-Steel Shot"
    CmbSAE.AddItem " - SAE J1993 - High-Carbon Cast-Steel Grit"
    
    'adicionar dados para o campo malha 1
    Lin = 5
    Do Until PlanBancoDeDados.Cells(Lin, 51) = ""
         CmbMalha1.AddItem PlanBancoDeDados.Cells(Lin, 51)
         Lin = Lin + 1
     Loop
     
   'adicionar dados para o campo malha 2
    Lin = 5
    Do Until PlanBancoDeDados.Cells(Lin, 51) = ""
         CmbMalha2.AddItem PlanBancoDeDados.Cells(Lin, 51)
         Lin = Lin + 1
    Loop
    
    'adicionar dados para o campo malha 3
    Lin = 5
    Do Until PlanBancoDeDados.Cells(Lin, 51) = ""
         CmbMalha3.AddItem PlanBancoDeDados.Cells(Lin, 51)
         Lin = Lin + 1
    Loop
         
    'adicionar dados para o campo malha 4
    Lin = 5
    Do Until PlanBancoDeDados.Cells(Lin, 51) = ""
         CmbMalha04.AddItem PlanBancoDeDados.Cells(Lin, 51)
         Lin = Lin + 1
    Loop
         
    'adicionar dados para o campo malha 5
    Lin = 5
    Do Until PlanBancoDeDados.Cells(Lin, 51) = ""
         CmbMalha5.AddItem PlanBancoDeDados.Cells(Lin, 51)
         Lin = Lin + 1
    Loop
         
    'adicionar dados para o campo malha 6
    Lin = 5
    Do Until PlanBancoDeDados.Cells(Lin, 51) = ""
         CmbMalha6.AddItem PlanBancoDeDados.Cells(Lin, 51)
         Lin = Lin + 1
    Loop
         
    'adicionar dados para o campo malha 7
    Lin = 5
    Do Until PlanBancoDeDados.Cells(Lin, 51) = ""
         CmbMalha7.AddItem PlanBancoDeDados.Cells(Lin, 51)
         Lin = Lin + 1
    Loop
    
    'adicionar dados para o campo 1 da especificação da granulometria
    CmbMáxMín1.AddItem "Máx. "
    CmbMáxMín1.AddItem "Mín. "
    
    'adicionar dados para o campo 2 da e´pecificação da granulometria
    CmbMáxMín2.AddItem "Máx. "
    CmbMáxMín2.AddItem "Mín. "
    
    'adicionar dados para o campo 3 da e´pecificação da granulometria
    CmbMáxMín3.AddItem "Máx. "
    CmbMáxMín3.AddItem "Mín. "
    
    'adicionar dados para o campo 4 da e´pecificação da granulometria
    CmbMáxMín4.AddItem "Máx. "
    CmbMáxMín4.AddItem "Mín. "
    
    'adicionar dados para o campo 5 da e´pecificação da granulometria
    CmbMáxMín5.AddItem "Máx. "
    CmbMáxMín5.AddItem "Mín. "
    
    'adicionar dados para o campo 6 da e´pecificação da granulometria
    CmbMáxMín6.AddItem "Máx. "
    CmbMáxMín6.AddItem "Mín. "
    
    'adicionar dados para o campo 7 da e´pecificação da granulometria
    CmbMáxMín7.AddItem "Máx. "
    CmbMáxMín7.AddItem "Mín. "
    
    'adicionar dados para o campo padrão
    cmbPadrão.AddItem "Shot"
    cmbPadrão.AddItem "Grit"
    
    'adicionar dados para o campo da microestrutura
    CmbMicro.AddItem "Martensita Revenida"
    CmbMicro.AddItem "Martensita"
    
    'adicionar dados para o campo quantidade de malhas
    cmbQuantMalhas.AddItem "2"
    cmbQuantMalhas.AddItem "3"
    cmbQuantMalhas.AddItem "4"
    cmbQuantMalhas.AddItem "5"
    cmbQuantMalhas.AddItem "6"
    cmbQuantMalhas.AddItem "7"
    
    'Deixar os campos de granulometria ocultos quando a macro é executada
    CmbMalha7.Visible = False
    CmbMáxMín7.Visible = False
    txtPorcentagem7.Visible = False
    CmbMalha6.Visible = False
    CmbMáxMín6.Visible = False
    txtPorcentagem6.Visible = False
    CmbMalha5.Visible = False
    CmbMáxMín5.Visible = False
    txtPorcentagem5.Visible = False
    CmbMalha04.Visible = False
    CmbMáxMín4.Visible = False
    txtPorcentagem4.Visible = False
    CmbMalha3.Visible = False
    CmbMáxMín3.Visible = False
    txtPorcentagem3.Visible = False
    CmbMalha2.Visible = False
    CmbMáxMín2.Visible = False
    txtPorcentagem2.Visible = False
    CmbMalha1.Visible = False
    CmbMáxMín1.Visible = False
    txtPorcentagem1.Visible = False
    
End Sub
```

•	Códigos que serão executados ao carregar o formulário

```
Private Sub cmbQuantMalhas_Change()
            
    'Se quantidade de malhas for igual a 2 mostrar 2 dados para especificação granulometrica
    If cmbQuantMalhas = "2" Then
            
    CmbMalha7.Visible = False
    CmbMáxMín7.Visible = False
    txtPorcentagem7.Visible = False
    CmbMalha6.Visible = False
    CmbMáxMín6.Visible = False
    txtPorcentagem6.Visible = False
    CmbMalha5.Visible = False
    CmbMáxMín5.Visible = False
    txtPorcentagem5.Visible = False
    CmbMalha04.Visible = False
    CmbMáxMín4.Visible = False
    txtPorcentagem4.Visible = False
    CmbMalha3.Visible = False
    CmbMáxMín3.Visible = False
    txtPorcentagem3.Visible = False
    CmbMalha2.Visible = True
    CmbMáxMín2.Visible = True
    txtPorcentagem2.Visible = True
    CmbMalha1.Visible = True
    CmbMáxMín1.Visible = True
    txtPorcentagem1.Visible = True
        
    'Se quantidade de malhas for igual a 3 mostrar 3 dados para especificação granulometrica
    
        ElseIf cmbQuantMalhas = "3" Then
            
        CmbMalha7.Visible = False
        CmbMáxMín7.Visible = False
        txtPorcentagem7.Visible = False
        CmbMalha6.Visible = False
        CmbMáxMín6.Visible = False
        txtPorcentagem6.Visible = False
        CmbMalha5.Visible = False
        CmbMáxMín5.Visible = False
        txtPorcentagem5.Visible = False
        CmbMalha04.Visible = False
        CmbMáxMín4.Visible = False
        txtPorcentagem4.Visible = False
        CmbMalha3.Visible = True
        CmbMáxMín3.Visible = True
        txtPorcentagem3.Visible = True
        CmbMalha2.Visible = True
        CmbMáxMín2.Visible = True
        txtPorcentagem2.Visible = True
        CmbMalha1.Visible = True
        CmbMáxMín1.Visible = True
        txtPorcentagem1.Visible = True
                    
    'Se quantidade de malhas for igual a 4 mostrar 4 dados para especificação granulometrica
    
            ElseIf cmbQuantMalhas = "4" Then
                
            CmbMalha7.Visible = False
            CmbMáxMín7.Visible = False
            txtPorcentagem7.Visible = False
            CmbMalha6.Visible = False
            CmbMáxMín6.Visible = False
            txtPorcentagem6.Visible = False
            CmbMalha5.Visible = False
            CmbMáxMín5.Visible = False
            txtPorcentagem5.Visible = False
            CmbMalha04.Visible = True
            CmbMáxMín4.Visible = True
            txtPorcentagem4.Visible = True
            CmbMalha3.Visible = True
            CmbMáxMín3.Visible = True
            txtPorcentagem3.Visible = True
            CmbMalha2.Visible = True
            CmbMáxMín2.Visible = True
            txtPorcentagem2.Visible = True
            CmbMalha1.Visible = True
            CmbMáxMín1.Visible = True
            txtPorcentagem1.Visible = True
                        
            'Se quantidade de malhas for igual a 5 mostrar 5 dados para especificação granulometrica
        
                ElseIf cmbQuantMalhas = "5" Then
                    
                CmbMalha7.Visible = False
                CmbMáxMín7.Visible = False
                txtPorcentagem7.Visible = False
                CmbMalha6.Visible = False
                CmbMáxMín6.Visible = False
                txtPorcentagem6.Visible = False
                CmbMalha5.Visible = True
                CmbMáxMín5.Visible = True
                txtPorcentagem5.Visible = True
                CmbMalha04.Visible = True
                CmbMáxMín4.Visible = True
                txtPorcentagem4.Visible = True
                CmbMalha3.Visible = True
                CmbMáxMín3.Visible = True
                txtPorcentagem3.Visible = True
                CmbMalha2.Visible = True
                CmbMáxMín2.Visible = True
                txtPorcentagem2.Visible = True
                CmbMalha1.Visible = True
                CmbMáxMín1.Visible = True
                txtPorcentagem1.Visible = True
                            
            'Se quantidade de malhas for igual a 6 mostrar 6 dados para especificação granulometrica
        
                    ElseIf cmbQuantMalhas = "6" Then
                        
                    CmbMalha7.Visible = False
                    CmbMáxMín7.Visible = False
                    txtPorcentagem7.Visible = False
                    CmbMalha6.Visible = True
                    CmbMáxMín6.Visible = True
                    txtPorcentagem6.Visible = True
                    CmbMalha5.Visible = True
                    CmbMáxMín5.Visible = True
                    txtPorcentagem5.Visible = True
                    CmbMalha04.Visible = True
                    CmbMáxMín4.Visible = True
                    txtPorcentagem4.Visible = True
                    CmbMalha3.Visible = True
                    CmbMáxMín3.Visible = True
                    txtPorcentagem3.Visible = True
                    CmbMalha2.Visible = True
                    CmbMáxMín2.Visible = True
                    txtPorcentagem2.Visible = True
                    CmbMalha1.Visible = True
                    CmbMáxMín1.Visible = True
                    txtPorcentagem1.Visible = True
                                    
            'Se quantidade de malhas for igual a 7 mostrar 7 dados para especificação granulometrica
            
                        ElseIf cmbQuantMalhas = "7" Then
                            
                        CmbMalha7.Visible = True
                        CmbMáxMín7.Visible = True
                        txtPorcentagem7.Visible = True
                        CmbMalha6.Visible = True
                        CmbMáxMín6.Visible = True
                        txtPorcentagem6.Visible = True
                        CmbMalha5.Visible = True
                        CmbMáxMín5.Visible = True
                        txtPorcentagem5.Visible = True
                        CmbMalha04.Visible = True
                        CmbMáxMín4.Visible = True
                        txtPorcentagem4.Visible = True
                        CmbMalha3.Visible = True
                        CmbMáxMín3.Visible = True
                        txtPorcentagem3.Visible = True
                        CmbMalha2.Visible = True
                        CmbMáxMín2.Visible = True
                        txtPorcentagem2.Visible = True
                        CmbMalha1.Visible = True
                        CmbMáxMín1.Visible = True
                        txtPorcentagem1.Visible = True
                        
    End If

    End Sub

```

•	Códigos que serão executados ao clicar no botão "Cadastrar"

```

Private Sub fmfCadastrar_Click()
'Cria uma variavel como string
Dim Resposta As String

    'Se o nome do novo produto estiver em branco, mostrar uma mensagem para digitar
        If txtNome = "" Then
            MsgBox "Digite o nome do produto"
            txtNome.SetFocus
            Exit Sub
        End If
    
    'Se o campo do padrão estiver em branco, mostrar uma mensagem para digitar
        If cmbPadrão = "" Then
            MsgBox "Selecione o 'Padrão' desejado"
            cmbPadrão.SetFocus
            Exit Sub
        
    'Se o campo do padrão estiver com dados que não sejam "Shot" ou "Grit", mostrar mensagem de valor inválido
        ElseIf cmbPadrão <> "Shot" And cmbPadrão <> "Grit" Then
            MsgBox "Digite um valor válido no campo 'Padrão'"
            cmbPadrão.SetFocus
            cmbPadrão = ""
            Exit Sub
        End If
            
        'Se o campo da composição química estiver em branco, mostrar uma mensagem para digitar
        If CmbCompQuím = "" Then
        MsgBox "Selecione a norma da composição química desejada"
        CmbCompQuím.SetFocus
        Exit Sub
    
    'Se o campo da composição química estiver com dados que não sejam as normas J827 ou J1993, mostrar mensagem de valor inválido
        ElseIf CmbCompQuím <> "Composição Química (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)" And CmbCompQuím <> "Composição Química (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)" Then
            MsgBox "Digite um valor válido no campo 'composição química'"
            CmbCompQuím.SetFocus
            CmbCompQuím = ""
            Exit Sub
        End If
        
            
    'Se o campo da dureza estiver em branco, mostrar uma mensagem para digitar
    If CmbDureza = "" Then
    MsgBox "Seleciona a norma da dureza desejada'"
    CmbDureza.SetFocus
    Exit Sub
    
    'Se o campo da dureza estiver com dados diferentes das normas J827, J1993 e BRS, mostrar mensagem de valor inválido
        ElseIf CmbDureza <> "Dureza (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)" And CmbDureza <> "Dureza (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)" And CmbDureza <> "Dureza (Baseado na norma BRS)" Then
        MsgBox "Digite um valor válido no campo 'dureza'"
        CmbDureza.SetFocus
        CmbDureza = ""
        Exit Sub
        End If
          
    'Se o campo da granulometria estiver em branco, mostrar uma mensagem para digitar
    If CmbGranulometria = "" Then
    MsgBox "Seleciona a norma da granulometria desejada'"
    CmbGranulometria.SetFocus
    Exit Sub
    
    'Se o campo da granulometria estiver com dados diferentes das normas J444 e BRS, mostrar mensagem de valor inválido
        ElseIf CmbGranulometria <> "Granulometria (Baseado na norma J444  - Cast Shot and Grit Size Specifications for Peening and Cleaning)" And CmbGranulometria <> "Granulometria (Baseado na norma BRS)" Then
        MsgBox "Digite um valor válido no campo 'granulometria'"
        CmbGranulometria.SetFocus
        CmbGranulometria = ""
        Exit Sub
        End If
        
    'Se o campo da densidade estiver em branco, mostrar uma mensagem para digitar
    If CmbDensidade = "" Then
    MsgBox "Seleciona a norma da densidade desejada'"
    CmbDensidade.SetFocus
    Exit Sub
    
    'Se o campo da densidade estiver com dados diferentes das normas J827, J1993 e BRS, mostrar mensagem de valor inválido
        ElseIf CmbDensidade <> "Densidade (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)" And CmbDensidade <> "Densidade (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)" And CmbDensidade <> "Densidade (Baseado na norma BRS)" Then
        MsgBox "Digite um valor válido no campo 'densidade'"
        CmbDensidade.SetFocus
        CmbDensidade = ""
        Exit Sub
        End If
     
    'Se o campo da macroestrutura estiver em branco, mostrar uma mensagem para digitar
    If CmbMacroestrutura = "" Then
    MsgBox "Seleciona a norma da macroestrutura desejada'"
    CmbMacroestrutura.SetFocus
    Exit Sub
    
    'Se o campo da macroestrutura estiver com dados diferentes das normas J827 e J1993 mostrar mensagem de valor inválido
        ElseIf CmbMacroestrutura <> "Macroestrutura (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)" And CmbMacroestrutura <> "Macroestrutura (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)" Then
        MsgBox "Digite um valor válido no campo 'macroestrutura"
        CmbMacroestrutura.SetFocus
        CmbMacroestrutura = ""
        Exit Sub
        End If
        
   'Se o campo da microestrutura estiver em branco, mostrar uma mensagem para digitar
    If CmbMicroestrutura = "" Then
    MsgBox "Seleciona a norma da microestrutura desejada'"
    CmbMicroestrutura.SetFocus
    Exit Sub
    
    'Se o campo da microestrutura estiver com dados diferentes das normas J827 e J1993 mostrar mensagem de valor inválido
        ElseIf CmbMicroestrutura <> "Microestrutura (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)" And CmbMicroestrutura <> "Microestrutura (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)" Then
        MsgBox "Digite um valor válido no campo 'microestrutura'"
        CmbMicroestrutura.SetFocus
        CmbMicroestrutura = ""
        Exit Sub
        End If
    
    'Se o campo do não magnéticos estiver em branco, mostrar uma mensagem para digitar
    If CmbNãoMagnético = "" Then
    MsgBox "Seleciona a norma do não magnéticos desejado'"
    CmbNãoMagnético.SetFocus
    Exit Sub
    
    'Se o campo do não magnéticos estiver com dados diferentes das normas J827 e J1993 mostrar mensagem de valor inválido
        ElseIf CmbNãoMagnético <> "Não magnéticos (Baseado na norma SAE J827 - High-Carbon Cast-Steel Shot)" And CmbNãoMagnético <> "Não magnéticos (Baseado na norma SAE J1993 - High-Carbon Cast-Steel Grit)" Then
        MsgBox "Digite um valor válido no campo 'não magnéticos'"
        CmbMicroestrutura.SetFocus
        CmbMacroestrutura = ""
        Exit Sub
        End If
        
    'Se o campo da SAE estver em branco, mostrar uma mensagem para digitar
    If CmbSAE = "" Then
    MsgBox "Seleciona a norma SAE desejada'"
    CmbSAE.SetFocus
    Exit Sub
    
    'Se o campo da SAE estiver com dados diferentes das normas J827 e J1993 mostrar mensagem de valor inválido
        ElseIf CmbSAE <> " - SAE J827 - High-Carbon Cast-Steel Shot" And CmbSAE <> " - SAE J1993 - High-Carbon Cast-Steel Grit" Then
        MsgBox "Digite um valor válido no campo 'SAE'"
        CmbSAE.SetFocus
        CmbSAE = ""
        Exit Sub
        End If
    
    'Se o primeiro campo da faixa de dureza estiver vazia, mostrar uma mensagem para digitar valor nela
    If TxtValordureza1 = "" Then
        MsgBox "Digite um valor para a faixa de dureza"
        TxtValordureza1.SetFocus
        Exit Sub
        
    'Se o primeiro campo da faixa de dureza não for númerico, mostrar uma mensagem de valor inválido
        ElseIf IsNumeric(TxtValordureza1) = False Then
            MsgBox "Valor inválido no campo de faixa de dureza"
            TxtValordureza1.Value = ""
            TxtValordureza1.SetFocus
            Exit Sub
    
    End If
    
     'Se o segundo campo da faixa de dureza estiver vazia, mostrar uma mensagem para digitar valor nela
    If TxtValorDureza2 = "" Then
        MsgBox "Digite um valor para a faixa de dureza"
        TxtValorDureza2.SetFocus
        Exit Sub
        
    'Se o segundo campo da faixa de dureza não for númerico, mostrar uma mensagem de valor inválido
        ElseIf IsNumeric(TxtValorDureza2) = False Then
            MsgBox "Valor inválido no campo de faixa de dureza"
            TxtValorDureza2.Value = ""
            TxtValorDureza2.SetFocus
            Exit Sub
    End If
    
    'Se a especificação da microestrutura estiver em branco, mostrar uma mensagem para digitar
    If CmbMicro = "" Then
    MsgBox "Seleciona a microestrutura desejada'"
    CmbMicro.SetFocus
    Exit Sub
    
    'Se o campo da microestrutura estiver com dados diferentes de Martensita e Martensita revenida mostrar uma mensagem
        ElseIf CmbMicro <> "Martensita" And CmbMicro <> "Martensita Revenida" Then
        MsgBox "Digite um valor válido no campo da especificação da microestrutura"
        CmbMicro.SetFocus
        CmbMicro = ""
        Exit Sub
    End If
    
    'Se o campo da especificação da densidade estiver em branco, mostrar uma mensagem para digitar
    If txtDensidade = "" Then
        MsgBox "Digite o valor mínimo da densidade"
        txtDensidade.SetFocus
        Exit Sub
    End If
    
    'Se o campo da especificação da % de esfera ou disforme estiver em branco, mostrar uma mensagem para digitar
    If txtEsfouDisf = "" Then
        MsgBox "Digite o valor máximo de esfera ou disforme"
        txtEsfouDisf.SetFocus
        Exit Sub
    End If
    
    'Analisa a quantidade de malhas seleciona
    Select Case cmbQuantMalhas
    
    'Caso for igual a 2, executa os comandos abaixo
    Case Is = "2"
        
    'Se o campo da primeira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha1 = "" Then
            MsgBox "Digite o valor da prmeira malha"
            CmbMalha1.SetFocus
            Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha1.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a primeira malha"
                CmbMalha1.SetFocus
                CmbMalha1 = ""
                Exit Sub
        End If
        
    'Se não for definido a primeira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín1 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na primeira malha"
        CmbMáxMín1.SetFocus
        Exit Sub
        
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín1.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín1.SetFocus
                CmbMáxMín1 = ""
                Exit Sub
        End If
        
    'Se o valor da primeira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem1 = "" Then
            MsgBox "Digite um valor para a primeira especificação granulometrica"
            txtPorcentagem1.SetFocus
            Exit Sub
        End If
        
     'Se o campo da segunda malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha2 = "" Then
            MsgBox "Digite o valor da segunda malha"
            CmbMalha2.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha2.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a segunda malha"
                CmbMalha2.SetFocus
                CmbMalha2 = ""
                Exit Sub
        End If
    
     'Se não for definido a segunda especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín2 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na segunda malha"
        CmbMáxMín2.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín2.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín2.SetFocus
                CmbMáxMín2 = ""
                Exit Sub
        End If
        
    'Se o valor da segunda especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem2 = "" Then
            MsgBox "Digite um valor para a segunda especificação granulometrica"
            txtPorcentagem2.SetFocus
            Exit Sub
        End If
        
    'Caso for igual a 3, executa os comandos abaixo
    Case Is = "3"
    
    'Se o campo da primeira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha1 = "" Then
            MsgBox "Digite o valor da prmeira malha"
            CmbMalha1.SetFocus
            Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha1.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a primeira malha"
                CmbMalha1.SetFocus
                CmbMalha1 = ""
                Exit Sub
        End If
        
    'Se não for definido a primeira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín1 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na primeira malha"
        CmbMáxMín1.SetFocus
        Exit Sub
        
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín1.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín1.SetFocus
                CmbMáxMín1 = ""
                Exit Sub
        End If
        
    'Se o valor da primeira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem1 = "" Then
            MsgBox "Digite um valor para a primeira especificação granulometrica"
            txtPorcentagem1.SetFocus
            Exit Sub
        End If
        
     'Se o campo da segunda malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha2 = "" Then
            MsgBox "Digite o valor da segunda malha"
            CmbMalha2.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha2.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a segunda malha"
                CmbMalha2.SetFocus
                CmbMalha2 = ""
                Exit Sub
        End If
    
     'Se não for definido a segunda especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín2 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na segunda malha"
        CmbMáxMín2.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín2.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín2.SetFocus
                CmbMáxMín2 = ""
                Exit Sub
        End If
        
    'Se o valor da segunda especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem2 = "" Then
            MsgBox "Digite um valor para a segunda especificação granulometrica"
            txtPorcentagem2.SetFocus
            Exit Sub
        End If
        
    'Se o campo da terceira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha3 = "" Then
            MsgBox "Digite o valor da terceira malha"
            CmbMalha3.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha3.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a terceira malha"
                CmbMalha3.SetFocus
                CmbMalha3 = ""
                Exit Sub
        End If
    
     'Se não for definido a terceira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín3 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na terceira malha"
        CmbMáxMín3.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín3.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín3.SetFocus
                CmbMáxMín3 = ""
                Exit Sub
        End If
        
    'Se o valor da terceira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem3 = "" Then
            MsgBox "Digite um valor para a terceira especificação granulometrica"
            txtPorcentagem3.SetFocus
            Exit Sub
        End If
        
'Caso for igual a 4, executa os comandos abaixo
    Case Is = "4"
    
    'Se o campo da primeira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha1 = "" Then
            MsgBox "Digite o valor da prmeira malha"
            CmbMalha1.SetFocus
            Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha1.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a primeira malha"
                CmbMalha1.SetFocus
                CmbMalha1 = ""
                Exit Sub
        End If
        
    'Se não for definido a primeira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín1 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na primeira malha"
        CmbMáxMín1.SetFocus
        Exit Sub
        
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín1.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín1.SetFocus
                CmbMáxMín1 = ""
                Exit Sub
        End If
        
    'Se o valor da primeira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem1 = "" Then
            MsgBox "Digite um valor para a primeira especificação granulometrica"
            txtPorcentagem1.SetFocus
            Exit Sub
        End If
        
     'Se o campo da segunda malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha2 = "" Then
            MsgBox "Digite o valor da segunda malha"
            CmbMalha2.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha2.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a segunda malha"
                CmbMalha2.SetFocus
                CmbMalha2 = ""
                Exit Sub
        End If
    
     'Se não for definido a segunda especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín2 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na segunda malha"
        CmbMáxMín2.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín2.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín2.SetFocus
                CmbMáxMín2 = ""
                Exit Sub
        End If
        
    'Se o valor da segunda especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem2 = "" Then
            MsgBox "Digite um valor para a segunda especificação granulometrica"
            txtPorcentagem2.SetFocus
            Exit Sub
        End If
        
    'Se o campo da terceira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha3 = "" Then
            MsgBox "Digite o valor da terceira malha"
            CmbMalha3.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha3.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a terceira malha"
                CmbMalha3.SetFocus
                CmbMalha3 = ""
                Exit Sub
        End If
    
     'Se não for definido a terceira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín3 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na terceira malha"
        CmbMáxMín3.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín3.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín3.SetFocus
                CmbMáxMín3 = ""
                Exit Sub
        End If
        
    'Se o valor da terceira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem3 = "" Then
            MsgBox "Digite um valor para a terceira especificação granulometrica"
            txtPorcentagem3.SetFocus
            Exit Sub
        End If
        
        'Se o campo da quarta malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha04 = "" Then
            MsgBox "Digite o valor da quarta malha"
            CmbMalha04.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha04.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a quarta malha"
                CmbMalha04.SetFocus
                CmbMalha04 = ""
                Exit Sub
        End If
    
     'Se não for definido a quarta especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín4 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na quarta malha"
        CmbMáxMín4.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín4.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín4.SetFocus
                CmbMáxMín4 = ""
                Exit Sub
        End If
        
    'Se o valor da quarta especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem4 = "" Then
            MsgBox "Digite um valor para a quarta especificação granulometrica"
            txtPorcentagem4.SetFocus
            Exit Sub
        End If
       
'Caso for igual a 5, executa os comandos abaixo
    Case Is = "5"
    
    'Se o campo da primeira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha1 = "" Then
            MsgBox "Digite o valor da prmeira malha"
            CmbMalha1.SetFocus
            Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha1.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a primeira malha"
                CmbMalha1.SetFocus
                CmbMalha1 = ""
                Exit Sub
        End If
        
    'Se não for definido a primeira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín1 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na primeira malha"
        CmbMáxMín1.SetFocus
        Exit Sub
        
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín1.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín1.SetFocus
                CmbMáxMín1 = ""
                Exit Sub
        End If
        
    'Se o valor da primeira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem1 = "" Then
            MsgBox "Digite um valor para a primeira especificação granulometrica"
            txtPorcentagem1.SetFocus
            Exit Sub
        End If
        
     'Se o campo da segunda malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha2 = "" Then
            MsgBox "Digite o valor da segunda malha"
            CmbMalha2.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha2.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a segunda malha"
                CmbMalha2.SetFocus
                CmbMalha2 = ""
                Exit Sub
        End If
    
     'Se não for definido a segunda especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín2 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na segunda malha"
        CmbMáxMín2.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín2.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín2.SetFocus
                CmbMáxMín2 = ""
                Exit Sub
        End If
        
    'Se o valor da segunda especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem2 = "" Then
            MsgBox "Digite um valor para a segunda especificação granulometrica"
            txtPorcentagem2.SetFocus
            Exit Sub
        End If
        
    'Se o campo da terceira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha3 = "" Then
            MsgBox "Digite o valor da terceira malha"
            CmbMalha3.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha3.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a terceira malha"
                CmbMalha3.SetFocus
                CmbMalha3 = ""
                Exit Sub
        End If
    
     'Se não for definido a terceira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín3 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na terceira malha"
        CmbMáxMín3.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín3.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín3.SetFocus
                CmbMáxMín3 = ""
                Exit Sub
        End If
        
    'Se o valor da terceira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem3 = "" Then
            MsgBox "Digite um valor para a terceira especificação granulometrica"
            txtPorcentagem3.SetFocus
            Exit Sub
        End If
        
        'Se o campo da quarta malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha04 = "" Then
            MsgBox "Digite o valor da quarta malha"
            CmbMalha04.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha04.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a quarta malha"
                CmbMalha04.SetFocus
                CmbMalha04 = ""
                Exit Sub
        End If
    
     'Se não for definido a quarta especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín4 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na quarta malha"
        CmbMáxMín4.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín4.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín4.SetFocus
                CmbMáxMín4 = ""
                Exit Sub
        End If
        
    'Se o valor da quarta especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem4 = "" Then
            MsgBox "Digite um valor para a quarta especificação granulometrica"
            txtPorcentagem4.SetFocus
            Exit Sub
        End If
        
        'Se o campo da quinta malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha5 = "" Then
            MsgBox "Digite o valor da quinta malha"
            CmbMalha5.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha5.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a quinta malha"
                CmbMalha5.SetFocus
                CmbMalha5 = ""
                Exit Sub
        End If
    
     'Se não for definido a quinta especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín5 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na quinta malha"
        CmbMáxMín5.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín5.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín5.SetFocus
                CmbMáxMín5 = ""
                Exit Sub
        End If
        
    'Se o valor da quinta especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem5 = "" Then
            MsgBox "Digite um valor para a quinta especificação granulometrica"
            txtPorcentagem5.SetFocus
            Exit Sub
        End If

'Caso for igual a 6, executa os comandos abaixo
    Case Is = "6"
    
    'Se o campo da primeira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha1 = "" Then
            MsgBox "Digite o valor da prmeira malha"
            CmbMalha1.SetFocus
            Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha1.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a primeira malha"
                CmbMalha1.SetFocus
                CmbMalha1 = ""
                Exit Sub
        End If
        
    'Se não for definido a primeira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín1 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na primeira malha"
        CmbMáxMín1.SetFocus
        Exit Sub
        
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín1.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín1.SetFocus
                CmbMáxMín1 = ""
                Exit Sub
        End If
        
    'Se o valor da primeira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem1 = "" Then
            MsgBox "Digite um valor para a primeira especificação granulometrica"
            txtPorcentagem1.SetFocus
            Exit Sub
        End If
        
     'Se o campo da segunda malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha2 = "" Then
            MsgBox "Digite o valor da segunda malha"
            CmbMalha2.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha2.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a segunda malha"
                CmbMalha2.SetFocus
                CmbMalha2 = ""
                Exit Sub
        End If
    
     'Se não for definido a segunda especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín2 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na segunda malha"
        CmbMáxMín2.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín2.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín2.SetFocus
                CmbMáxMín2 = ""
                Exit Sub
        End If
        
    'Se o valor da segunda especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem2 = "" Then
            MsgBox "Digite um valor para a segunda especificação granulometrica"
            txtPorcentagem2.SetFocus
            Exit Sub
        End If
        
    'Se o campo da terceira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha3 = "" Then
            MsgBox "Digite o valor da terceira malha"
            CmbMalha3.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha3.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a terceira malha"
                CmbMalha3.SetFocus
                CmbMalha3 = ""
                Exit Sub
        End If
    
     'Se não for definido a terceira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín3 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na terceira malha"
        CmbMáxMín3.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín3.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín3.SetFocus
                CmbMáxMín3 = ""
                Exit Sub
        End If
        
    'Se o valor da terceira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem3 = "" Then
            MsgBox "Digite um valor para a terceira especificação granulometrica"
            txtPorcentagem3.SetFocus
            Exit Sub
        End If
        
        'Se o campo da quarta malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha04 = "" Then
            MsgBox "Digite o valor da quarta malha"
            CmbMalha04.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha04.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a quarta malha"
                CmbMalha04.SetFocus
                CmbMalha04 = ""
                Exit Sub
        End If
    
     'Se não for definido a quarta especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín4 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na quarta malha"
        CmbMáxMín4.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín4.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín4.SetFocus
                CmbMáxMín4 = ""
                Exit Sub
        End If
        
    'Se o valor da quarta especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem4 = "" Then
            MsgBox "Digite um valor para a quarta especificação granulometrica"
            txtPorcentagem4.SetFocus
            Exit Sub
        End If
        
        'Se o campo da quinta malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha5 = "" Then
            MsgBox "Digite o valor da quinta malha"
            CmbMalha5.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha5.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a quinta malha"
                CmbMalha5.SetFocus
                CmbMalha5 = ""
                Exit Sub
        End If
    
     'Se não for definido a quinta especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín5 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na quinta malha"
        CmbMáxMín5.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín5.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín5.SetFocus
                CmbMáxMín5 = ""
                Exit Sub
        End If
        
    'Se o valor da quinta especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem5 = "" Then
            MsgBox "Digite um valor para a quinta especificação granulometrica"
            txtPorcentagem5.SetFocus
            Exit Sub
        End If
        
        'Se o campo da sexta malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha6 = "" Then
            MsgBox "Digite o valor da sexta malha"
            CmbMalha6.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha6.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a sexta malha"
                CmbMalha6.SetFocus
                CmbMalha6 = ""
                Exit Sub
        End If
    
     'Se não for definido a sexta especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín6 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na sexta malha"
        CmbMáxMín6.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín6.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín6.SetFocus
                CmbMáxMín6 = ""
                Exit Sub
        End If
        
    'Se o valor da sexta especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem6 = "" Then
            MsgBox "Digite um valor para a sexta especificação granulometrica"
            txtPorcentagem6.SetFocus
            Exit Sub
        End If

'Caso for igual a 7, executa os comandos abaixo
    Case Is = "7"
    
    'Se o campo da primeira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha1 = "" Then
            MsgBox "Digite o valor da prmeira malha"
            CmbMalha1.SetFocus
            Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha1.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a primeira malha"
                CmbMalha1.SetFocus
                CmbMalha1 = ""
                Exit Sub
        End If
        
    'Se não for definido a primeira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín1 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na primeira malha"
        CmbMáxMín1.SetFocus
        Exit Sub
        
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín1.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín1.SetFocus
                CmbMáxMín1 = ""
                Exit Sub
        End If
        
    'Se o valor da primeira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem1 = "" Then
            MsgBox "Digite um valor para a primeira especificação granulometrica"
            txtPorcentagem1.SetFocus
            Exit Sub
        End If
        
     'Se o campo da segunda malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha2 = "" Then
            MsgBox "Digite o valor da segunda malha"
            CmbMalha2.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha2.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a segunda malha"
                CmbMalha2.SetFocus
                CmbMalha2 = ""
                Exit Sub
        End If
    
     'Se não for definido a segunda especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín2 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na segunda malha"
        CmbMáxMín2.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín2.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín2.SetFocus
                CmbMáxMín2 = ""
                Exit Sub
        End If
        
    'Se o valor da segunda especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem2 = "" Then
            MsgBox "Digite um valor para a segunda especificação granulometrica"
            txtPorcentagem2.SetFocus
            Exit Sub
        End If
        
    'Se o campo da terceira malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha3 = "" Then
            MsgBox "Digite o valor da terceira malha"
            CmbMalha3.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha3.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a terceira malha"
                CmbMalha3.SetFocus
                CmbMalha3 = ""
                Exit Sub
        End If
    
     'Se não for definido a terceira especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín3 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na terceira malha"
        CmbMáxMín3.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín3.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín3.SetFocus
                CmbMáxMín3 = ""
                Exit Sub
        End If
        
    'Se o valor da terceira especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem3 = "" Then
            MsgBox "Digite um valor para a terceira especificação granulometrica"
            txtPorcentagem3.SetFocus
            Exit Sub
        End If
        
        'Se o campo da quarta malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha04 = "" Then
            MsgBox "Digite o valor da quarta malha"
            CmbMalha04.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha04.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a quarta malha"
                CmbMalha04.SetFocus
                CmbMalha04 = ""
                Exit Sub
        End If
    
     'Se não for definido a quarta especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín4 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na quarta malha"
        CmbMáxMín4.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín4.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín4.SetFocus
                CmbMáxMín4 = ""
                Exit Sub
        End If
        
    'Se o valor da quarta especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem4 = "" Then
            MsgBox "Digite um valor para a quarta especificação granulometrica"
            txtPorcentagem4.SetFocus
            Exit Sub
        End If
        
        'Se o campo da quinta malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha5 = "" Then
            MsgBox "Digite o valor da quinta malha"
            CmbMalha5.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha5.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a quinta malha"
                CmbMalha5.SetFocus
                CmbMalha5 = ""
                Exit Sub
        End If
    
     'Se não for definido a quinta especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín5 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na quinta malha"
        CmbMáxMín5.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín5.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín5.SetFocus
                CmbMáxMín5 = ""
                Exit Sub
        End If
        
    'Se o valor da quinta especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem5 = "" Then
            MsgBox "Digite um valor para a quinta especificação granulometrica"
            txtPorcentagem5.SetFocus
            Exit Sub
        End If
        
        'Se o campo da sexta malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha6 = "" Then
            MsgBox "Digite o valor da sexta malha"
            CmbMalha6.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha6.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a sexta malha"
                CmbMalha6.SetFocus
                CmbMalha6 = ""
                Exit Sub
        End If
    
     'Se não for definido a sexta especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín6 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na sexta malha"
        CmbMáxMín6.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín6.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín6.SetFocus
                CmbMáxMín6 = ""
                Exit Sub
        End If
        
    'Se o valor da sexta especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem6 = "" Then
            MsgBox "Digite um valor para a sexta especificação granulometrica"
            txtPorcentagem6.SetFocus
            Exit Sub
        End If

        'Se o campo da sétima malha estiver vazia, mstrar uma mensagem para digitar
        If CmbMalha7 = "" Then
            MsgBox "Digite o valor da sétima malha"
            CmbMalha7.SetFocus
            Exit Sub
            
    'Se o usuário digitar um valor que não esteja na lista de malhas, mostrar uma mensagem
                ElseIf CmbMalha7.ListIndex = -1 Then
                MsgBox "Escolha um valor dentro da lista de opções para a sétima malha"
                CmbMalha7.SetFocus
                CmbMalha7 = ""
                Exit Sub
        End If
    
     'Se não for definido a sétima especificação de máximo e mínimo, mostrar uma mensagem
        If CmbMáxMín7 = "" Then
        MsgBox "Defina se a especificação é máxima ou mínima na sétima malha"
        CmbMáxMín7.SetFocus
        Exit Sub
    
    'Se o usuário digitar um valor que não esteja na lista(máx ou Min), mostrar uma mensagem
                ElseIf CmbMáxMín7.ListIndex = -1 Then
                MsgBox "valor inválido, escolha se é máximo ou mínimo"
                CmbMáxMín7.SetFocus
                CmbMáxMín7 = ""
                Exit Sub
        End If
        
    'Se o valor da sétima especificação estiver vazia, mostrar uma mensagem
        If txtPorcentagem7 = "" Then
            MsgBox "Digite um valor para a sétima especificação granulometrica"
            txtPorcentagem7.SetFocus
            Exit Sub
        End If
        
 End Select
    
    'Transferir os dados do formulário para a planilha "Banco de dados" do excel
    PlanBancoDeDados.Select
    'Seleciona a célula abaixo do último campo em branco do banco de dados
    Range("A168").Select
    'O mesmo que aperta CTRL + SETA PRA CIMA, sobe de célula em célula até encontrar uma preenchimento
    ActiveCell.End(xlUp).Select
    'Descloca a célula ativa para uma linha para baixo
    ActiveCell.Offset(1, 0).Select
    'Tranfere o dado inserido no nome do userform para o banco de dados excel
    ActiveCell.Value = txtNome
    'Tranfere o dado inserido na norma composição química do userform para o banco de dados excel
    ActiveCell.Offset(0, 3).Value = CmbCompQuím
    'Tranfere o dado inserido na norma dureza do userform para o banco de dados excel
    ActiveCell.Offset(0, 4).Value = CmbDureza
    'Tranfere o dado inserido na norma granulometria do userform para o banco de dados excel
    ActiveCell.Offset(0, 5).Value = CmbGranulometria
    'Tranfere o dado inserido na norma densidade do userform para o banco de dados excel
    ActiveCell.Offset(0, 6).Value = CmbDensidade
    'Tranfere o dado inserido na norma macroestrutura do userform para o banco de dados excel
    ActiveCell.Offset(0, 7).Value = CmbMacroestrutura
    'Tranfere o dado inserido na norma microestrutura do userform para o banco de dados excel
    ActiveCell.Offset(0, 8).Value = CmbMicroestrutura
    'Tranfere o dado inserido na norma não magnéticos do userform para o banco de dados excel
    ActiveCell.Offset(0, 9).Value = CmbNãoMagnético
    'Tranfere o dado inserido na norma SAE do userform para o banco de dados excel
    ActiveCell.Offset(0, 10).Value = CmbSAE
    'Tranfere o dado inserido no valor mínimo da faixa de dureza para o banco de dados excel
    ActiveCell.Offset(0, 42).Value = TxtValordureza1
    'Tranfere o dado inserido no valor máximo da faixa de dureza para o banco de dados excel
    ActiveCell.Offset(0, 43).Value = TxtValorDureza2
    'Tranfere o dado inserido na primeira malha do userform para o banco de dados excel
    ActiveCell.Offset(0, 13).Value = CmbMalha1
    'Tranfere o dado inserido na primeira  malha do máx ou mín do userform para o banco de dados excel
    ActiveCell.Offset(0, 44).Value = CmbMáxMín1
    'Tranfere o dado inserido na primeira  malha da especificação do userform para o banco de dados excel
    ActiveCell.Offset(0, 45).Value = txtPorcentagem1
    'Tranfere o dado inserido na segunda malha do userform para o banco de dados excel
    ActiveCell.Offset(0, 16).Value = CmbMalha2
    'Tranfere o dado inserido na segunda  malha do máx ou mín do userform para o banco de dados excel
    ActiveCell.Offset(0, 46).Value = CmbMáxMín2
    'Tranfere o dado inserido na segunda  malha da especificação do userform para o banco de dados excel
    ActiveCell.Offset(0, 47).Value = txtPorcentagem2
    'Tranfere o dado inserido na terceira malha do userform para o banco de dados excel
    ActiveCell.Offset(0, 19).Value = CmbMalha3
    'Tranfere o dado inserido na terceira  malha do máx ou mín do userform para o banco de dados excel
    ActiveCell.Offset(0, 48).Value = CmbMáxMín3
    'Tranfere o dado inserido na terceira  malha da especificação do userform para o banco de dados excel
    ActiveCell.Offset(0, 49).Value = txtPorcentagem3
    'Tranfere o dado inserido na quarta malha do userform para o banco de dados excel
    ActiveCell.Offset(0, 22).Value = CmbMalha04
    'Tranfere o dado inserido na quarta  malha do máx ou mín do userform para o banco de dados excel
    ActiveCell.Offset(0, 50).Value = CmbMáxMín4
    'Tranfere o dado inserido na quarta  malha da especificação do userform para o banco de dados excel
    ActiveCell.Offset(0, 51).Value = txtPorcentagem4
    'Tranfere o dado inserido na quinta malha do userform para o banco de dados excel
    ActiveCell.Offset(0, 25).Value = CmbMalha5
    'Tranfere o dado inserido na quinta  malha do máx ou mín do userform para o banco de dados excel
    ActiveCell.Offset(0, 52).Value = CmbMáxMín5
    'Tranfere o dado inserido na quinta  malha da especificação do userform para o banco de dados excel
    ActiveCell.Offset(0, 53).Value = txtPorcentagem5
    'Tranfere o dado inserido na sexta malha do userform para o banco de dados excel
    ActiveCell.Offset(0, 28).Value = CmbMalha6
    'Tranfere o dado inserido na sexta  malha do máx ou mín do userform para o banco de dados excel
    ActiveCell.Offset(0, 54).Value = CmbMáxMín6
    'Tranfere o dado inserido na sexta  malha da especificação do userform para o banco de dados excel
    ActiveCell.Offset(0, 55).Value = txtPorcentagem6
    'Tranfere o dado inserido na sétima malha do userform para o banco de dados excel
    ActiveCell.Offset(0, 31).Value = CmbMalha7
    'Tranfere o dado inserido na sétima  malha do máx ou mín do userform para o banco de dados excel
    ActiveCell.Offset(0, 56).Value = CmbMáxMín7
    'Tranfere o dado inserido na sétima  malha da especificação do userform para o banco de dados excel
    ActiveCell.Offset(0, 57).Value = txtPorcentagem7
    'Tranfere o dado inserido na densidade do userform para o banco de dados do excel
    ActiveCell.Offset(0, 58).Value = txtDensidade
    'Tranfere o dado inserido na %esfera ou disforme do userform para o banco de dados do excel
    ActiveCell.Offset(0, 59).Value = txtEsfouDisf
    'Tranfere o dado inserido na especifacação da microestrutura do userform para o banco de dados do excel
    ActiveCell.Offset(0, 35).Value = CmbMicro
    'Tranfere o dado inserido no padrão da userform para o banco de dados do excel
    ActiveCell.Offset(0, 37).Value = cmbPadrão
    
    'Seleciona a planilha 1 ("Ficha técnica)
    Plan1.Select
    
    'mostra uma mensagem para o usuário
    MsgBox "Material cadastrado com sucesso", vbExclamation, "Ficha tecnica"
    
    'Cria uma variável para salientar se o usuário quer emitir a ficha técnica criada
    Resposta = MsgBox("Gostaria de emitir essa ficha tecnica ?", vbYesNo, "Ficha tecnica")
    
    'Se a resposta for sim
    If Resposta = vbYes Then
    Unload Me
    frmEmitir.CmbIdioma = ""
    frmEmitir.CmbMaterial = ""
    
    'Se a resposta for não
    ElseIf Resposta = vbNo Then
    Unload Me
    Unload frmEmitir
    End If
    
End Sub

```
8. Por fim, cria-se códigos para rodar quando abre o documento excel

```
Private Sub Workbook_Open()

'Seleciona a planilha 1 (Modelo)
Plan1.Select

'True =  ativa o excel com a tela cheia, False = mantem a tela minimizada
Application.DisplayFullScreen = False

'True = ativa a barra de fórmula / False = oculta a barra de fórmula
Application.DisplayFormulaBar = True

'True = Cabeçalhos tanto de coluna quanto de linha serão exibidos /False = deixa-os ocultados
ActiveWindow.DisplayHeadings = False

'True = Linhas de grade estiverem exibidas / False = não serão exibidas as linhas de grade
ActiveWindow.DisplayGridlines = False

'true = Habilita a visualização das abas das planilhas / False = desabilita a visualização 
ActiveWindow.DisplayWorkbookTabs = False

'Senha para proteger e ninguem mexer no código
ActiveWorkbook.Protect "XXXXXX"

End Sub

```
