'Apontamento de Consumo de Concreto

Private Sub UserForm_Activate()

    ' Verificar se a variável global "usernameGlobal" está vazia
    If usernameGlobal = "" Then
        ' Fechar o formulário atual
        Me.Hide
        ' Redirecionar para o formulário de login
        Login.Show
    End If

    Dim leitura As Worksheet
    Set leitura = ThisWorkbook.Worksheets("validConcreto")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("recebConcreto")
    
    With ws
    Dim lastRow As Long
    lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        
    Dim i As Long
    Dim lista As String
        
    For i = lastRow To 2 Step -1 'começa do último item
        lista = lista & .Cells(i, 1).Value & vbCrLf
    Next i
        
    ListBox1.List = Split(lista, vbCrLf)
    End With
    
'lista os ultimos cadastros
    
    Dim ws0 As Worksheet
    Set ws0 = ThisWorkbook.Worksheets("consConcreto")

    With ws0
    Dim lastRow0 As Long
    lastRow0 = .Cells(.Rows.Count, 1).End(xlUp).Row

    Dim i0 As Long
    Dim lista0 As String

    For i0 = lastRow0 To 2 Step -1 'começa do último item
        lista0 = lista0 & .Cells(i0, 1).Value & vbCrLf
    Next i0

    ListBox2.Clear ' Limpa a lista atual do ListBox2
    ListBox2.List = Split(lista0, vbCrLf)
    End With

    
'Lista os destinos
    
    ComboBox4.AddItem "Parede"
    ComboBox4.AddItem "Pilar"
    ComboBox4.AddItem "Radier"
    ComboBox4.AddItem "Calçada"
    
'Fim da listagem de destinos

    ComboBox5.AddItem "Infraestrutura"
    ComboBox5.AddItem "Habitação"
    ComboBox5.AddItem "Comunitário"
    
    ComboBox6.AddItem "Bomba Estacionária"
    ComboBox6.AddItem "Bomba Lança"
    ComboBox6.AddItem "Convencional"
       
End Sub

Private Sub ComboBox5_Change()

    If ComboBox5.Value = "Habitação" Then
        ' Limpar ComboBox1 e ComboBox2
        ComboBox1.Clear
        ComboBox2.Clear
        
        ' Adicionar itens na ComboBox1
        ComboBox1.AddItem "A"
        ComboBox1.AddItem "B"
        ComboBox1.AddItem "C"
        ComboBox1.AddItem "D"
        ComboBox1.AddItem "E"
        ComboBox1.AddItem "F"
        ComboBox1.AddItem "G"
        ComboBox1.AddItem "H"
        ComboBox1.AddItem "I"
        ComboBox1.AddItem "J"
        ComboBox1.AddItem "K"
        ComboBox1.AddItem "L"
        ComboBox1.AddItem "M"
        ComboBox1.AddItem "N"
        ComboBox1.AddItem "O"
        ComboBox1.AddItem "P"
        ComboBox1.AddItem "Q"
        ComboBox1.AddItem "R"
        ComboBox1.AddItem "U"
        ComboBox1.AddItem "V"
        ComboBox1.AddItem "W"
        
        ' Adicionar itens na ComboBox2
        ComboBox2.AddItem "1"
        ComboBox2.AddItem "2"
        ComboBox2.AddItem "3"
        ComboBox2.AddItem "4"
        ComboBox2.AddItem "5"
        ComboBox2.AddItem "6"
        ComboBox2.AddItem "7"
        ComboBox2.AddItem "8"
        ComboBox2.AddItem "9"
        ComboBox2.AddItem "10"
        ComboBox2.AddItem "11"
        ComboBox2.AddItem "12"
        ComboBox2.AddItem "13"
        ComboBox2.AddItem "14"
        ComboBox2.AddItem "15"
        ComboBox2.AddItem "16"
        ComboBox2.AddItem "17"
        ComboBox2.AddItem "18"
        ComboBox2.AddItem "19"
        ComboBox2.AddItem "20"
        ComboBox2.AddItem "21"
        ComboBox2.AddItem "22"
        ComboBox2.AddItem "23"
        ComboBox2.AddItem "24"
        ComboBox2.AddItem "25"
        ComboBox2.AddItem "26"
        ComboBox2.AddItem "27"
        ComboBox2.AddItem "28"
        ComboBox2.AddItem "29"
        ComboBox2.AddItem "30"
        ComboBox2.AddItem "31"
        ComboBox2.AddItem "32"
        ComboBox2.AddItem "33"
        ComboBox2.AddItem "34"
        ComboBox2.AddItem "35"
        ComboBox2.AddItem "36"
        ComboBox2.AddItem "37"
        ComboBox2.AddItem "38"
        ComboBox2.AddItem "39"
        ComboBox2.AddItem "40"
        ComboBox2.AddItem "41"
        ComboBox2.AddItem "42"
        ComboBox2.AddItem "43"
        ComboBox2.AddItem "44"
        ComboBox2.AddItem "45"
        ComboBox2.AddItem "46"
        ComboBox2.AddItem "47"

    ElseIf ComboBox5.Value = "Infraestrutura" Then
        ' Limpar ComboBox1 e ComboBox2
        ComboBox1.Clear
        ComboBox2.Clear
        ComboBox1.Value = "Geral"
        ComboBox2.Value = "Geral"
    
    ElseIf ComboBox5.Value = "Comunitário" Then
        ' Limpar ComboBox1 e ComboBox2
        ComboBox1.Clear
        ComboBox2.Clear
        ComboBox1.Value = "Geral"
        ComboBox2.Value = "Geral"
    End If
    
    ' Realizar busca de valores não repetidos na coluna L com base no valor da coluna O
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim uniqueValues As Collection
    Dim cell As Range
    
    Set ws = ThisWorkbook.Worksheets("validConcreto")
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    Set dataRange = ws.Range("L2:L" & lastRow)
    Set uniqueValues = New Collection
    
    On Error Resume Next
    For Each cell In dataRange
        If cell.Offset(0, 3).Value = ComboBox5.Value Then ' Coluna O
            uniqueValues.Add cell.Value, CStr(cell.Value)
        End If
    Next cell
    On Error GoTo 0
    
    ' Preencher ComboBox4 com os valores não repetidos encontrados
    ComboBox4.Clear
    For Each Value In uniqueValues
        ComboBox4.AddItem Value
    Next Value
    
End Sub

Private Sub TextBox17_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii >= 48 And KeyAscii <= 57 Then ' permite apenas números de 0 a 9
        If Len(Me.TextBox17.Value) = 2 Then ' adiciona dois pontos após a hora ser digitada
            Me.TextBox17.Value = Me.TextBox17.Value & ":"
            Me.TextBox17.SelStart = Len(Me.TextBox17.Value)
        ElseIf Len(Me.TextBox17.Value) = 5 Then ' limita a entrada em 5 caracteres para evitar que o usuário digite mais do que HH:MM
            Me.TextBox17.Value = Me.TextBox17.Value & ":"
            Me.TextBox17.SelStart = Len(Me.TextBox17.Value)
        End If
    ElseIf KeyAscii = 8 Then ' permite uso da tecla backspace
        If Len(Me.TextBox17.Value) = 3 Then ' remove o caracter ":" se ele existir
            Me.TextBox17.Value = Left(Me.TextBox17.Value, 2)
            Me.TextBox17.SelStart = Len(Me.TextBox17.Value)
        ElseIf Len(Me.TextBox17.Value) = 6 Then ' remove o caracter ":" se ele existir
            Me.TextBox17.Value = Left(Me.TextBox17.Value, 5)
            Me.TextBox17.SelStart = Len(Me.TextBox17.Value)
        End If
    Else
        KeyAscii = 0 ' desabilita a tecla pressionada caso não seja um número ou backspace
    End If
End Sub

Private Sub ListBox1_AfterUpdate()
    ' Verifica se há uma linha selecionada na ListBox1
    If ListBox1.ListIndex <> -1 Then
        ' Obtém o valor da linha selecionada na ListBox1
        Dim selectedValue As String
        selectedValue = ListBox1.Value
        
        ' Define a planilha onde os dados estão armazenados
        Dim wsCons As Worksheet
        Dim wsReceb As Worksheet
        Set wsCons = ThisWorkbook.Worksheets("consConcreto")
        Set wsReceb = ThisWorkbook.Worksheets("recebConcreto")
        
        ' Busca os valores correspondentes nas colunas D, E, F e G da aba "consConcreto"
        Dim consLastRow As Long
        consLastRow = wsCons.Cells(wsCons.Rows.Count, "C").End(xlUp).Row
        Dim consTotal As Double
        Dim consValue As String
        Dim i As Long
        For i = 2 To consLastRow
            If wsCons.Cells(i, "C").Value = selectedValue Then
                consTotal = consTotal + CDbl(wsCons.Cells(i, "J").Value)
                consValue = wsCons.Cells(i, "G").Value
            End If
        Next i
        
        ' Busca o valor correspondente na coluna A da aba "recebConcreto"
        Dim recebLastRow As Long
        recebLastRow = wsReceb.Cells(wsReceb.Rows.Count, "A").End(xlUp).Row
        Dim recebValue As Double
        Dim recebTextBox4 As String
        For i = 2 To recebLastRow
            If wsReceb.Cells(i, "A").Value = selectedValue Then
                recebValue = CDbl(wsReceb.Cells(i, "O").Value)
                recebTextBox4 = wsReceb.Cells(i, "G").Value
                Exit For
            End If
        Next i
        
        ' Calcula a diferença entre os valores
        Dim diferenca As Double
        diferenca = recebValue - consTotal
        
        ' Exibe os resultados nas TextBoxes correspondentes
        TextBox1.Value = wsReceb.Cells(i, "D").Value
        TextBox2.Value = wsReceb.Cells(i, "E").Value
        TextBox3.Value = wsReceb.Cells(i, "F").Value
        TextBox4.Value = recebTextBox4
        TextBox26.Value = wsReceb.Cells(i, "N").Value
        TextBox5.Value = wsReceb.Cells(i, "H").Value
        TextBox24.Value = recebValue
        TextBox25.Value = diferenca
    End If
End Sub

Private Sub ListBox2_AfterUpdate()
    ' Verifica se há uma linha selecionada na ListBox2
    If ListBox2.ListIndex <> -1 Then
        ' Obtém o valor da linha selecionada na ListBox2
        Dim selectedValue As String
        selectedValue = ListBox2.Value
        
        ' Define a planilha onde os dados estão armazenados
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("consConcreto")
        
        ' Procura o valor na coluna A da planilha e retorna a linha correspondente
        Dim searchRange As Range
        Set searchRange = ws.Columns("A")
        
        Dim foundCell As Range
        Set foundCell = searchRange.Find(What:=selectedValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Verifica se o valor foi encontrado
        If Not foundCell Is Nothing Then
            ' Obtém a linha correspondente ao valor encontrado
            Dim selectedRow As Long
            selectedRow = foundCell.Row
            
            ' Define o valor dos TextBoxes como os valores das colunas correspondentes da linha selecionada
            TextBox18.Value = Format(ws.Cells(selectedRow, "G").Value, "hh:mm:ss")
            TextBox19.Value = ws.Cells(selectedRow, "F").Value
            TextBox20.Value = ws.Cells(selectedRow, "H").Value
            TextBox21.Value = ws.Cells(selectedRow, "I").Value
            TextBox22.Value = ws.Cells(selectedRow, "J").Value
            TextBox23.Value = ws.Cells(selectedRow, "K").Value
            TextBox24.Value = Format(ws.Cells(selectedRow, "O").Value, "#0.00")
        End If
    End If
End Sub

Private Sub ComboBox2_Change()
    If ComboBox5.Value = "Habitação" Then
        Dim ws As Worksheet
        Dim searchValue As String
        Dim concatValue As String
        Dim foundCell As Range
        
        Set ws = ThisWorkbook.Worksheets("validConcreto")
        
        searchValue = Me.ComboBox2.Value
        concatValue = Me.ComboBox1.Value & searchValue
        
        Set foundCell = ws.Columns("F:F").Find(What:=concatValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            Me.TextBox15.Value = foundCell.Offset(0, 3).Value ' Coluna I
        Else
            MsgBox "A unidade não foi encontrada na previsão, por favor, indique uma unidade existente.", vbExclamation
            Me.TextBox15.Value = ""
            Me.ComboBox1.Value = ""
            Me.ComboBox2.Value = ""
            ComboBox1.SetFocus
        End If
    Else
        
        Me.TextBox15.Value = "Único"
        
    End If
End Sub

Private Sub ComboBox4_Change()
    Dim ws As Worksheet
    Dim searchValue As String
    Dim concatValue As String
    Dim foundCell As Range
    
    Set ws = ThisWorkbook.Worksheets("validConcreto")
    
    searchValue = Me.ComboBox4.Value
    concatValue = Me.TextBox15.Value & searchValue
    
    Set foundCell = ws.Columns("K:K").Find(What:=concatValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        Me.TextBox16.Value = foundCell.Offset(0, 3).Value ' Coluna N
    Else
        Me.TextBox16.Value = ""
    End If
End Sub

Private Sub CommandButton1_Click()
    ' Verifica se todos os campos foram preenchidos
    If TextBox1.Value = "" Or ComboBox1.Value = "" Or ComboBox2.Value = "" _
        Or TextBox26.Value = "" Or ComboBox4.Value = "" Or TextBox13.Value = "" _
        Or TextBox14.Value = "" Then
        MsgBox "Por favor, preencha todos os campos.", vbExclamation, "Campos obrigatórios"
        Exit Sub
    End If
    
    ' Verifica se TextBox13 contém um número decimal válido
    If Not IsNumeric(TextBox13.Value) Then
        MsgBox "Por favor, preencha um número decimal válido no valor de bombeamento executado.", vbExclamation, "Número inválido"
        TextBox13.SetFocus
        Exit Sub
    End If
    
    ' Verifica se TextBox14 contém um número decimal válido
    If Not IsNumeric(TextBox14.Value) Then
        MsgBox "Por favor, preencha um número decimal válido no valor de bombeamento medido.", vbExclamation, "Número inválido"
        TextBox14.SetFocus
        Exit Sub
    End If
    
    ' Verifica se TextBox17 contém um horário válido no formato HH:MM
    If Not IsDate(TextBox17.Value) Then
        MsgBox "Por favor, preencha um horário válido no formato HH:MM no campo Fim do Bombeamento.", vbExclamation, "Campo inválido"
        TextBox17.SetFocus
        Exit Sub
    End If
    
 ' Verifica se o valor de TextBox13 é superior ao valor de TextBox25
    If CDbl(TextBox13.Value) > CDbl(TextBox25.Value) Then
        MsgBox "A quantidade inserida excede o valor restante do caminhão, por favor, insira quantidade válida.", vbExclamation, "Quantidade inválida"
        Exit Sub
    End If
    
    
    If CDbl(TextBox13.Value) > CDbl(TextBox16.Value) And CDbl(TextBox16.Value) <> 0 Then
    Dim observacao As String
    observacao = InputBox("O valor executado é superior ao previsto para o destino informado. Por favor, insira uma observação:", "Superior ao previsto")

    ' Verifica se o usuário cancelou a inserção da observação
    If observacao = "" Then
        MsgBox "Cadastro cancelado pelo usuário.", vbInformation, "Cancelado"
        Exit Sub
    End If

    ' Procura o valor na coluna A da planilha recebConcreto e retorna a linha correspondente
    Dim wsReceb As Worksheet
    Set wsReceb = ThisWorkbook.Worksheets("recebConcreto")

    Dim searchRangeReceb As Range
    Set searchRangeReceb = wsReceb.Columns("A")

    Dim foundCellReceb As Range
    Set foundCellReceb = searchRangeReceb.Find(What:=ListBox1.Value, LookIn:=xlValues, LookAt:=xlWhole)

    ' Verifica se o valor foi encontrado na planilha recebConcreto
    If Not foundCellReceb Is Nothing Then
        ' Obtém a linha correspondente ao valor encontrado
        Dim selectedRowReceb As Long
        selectedRowReceb = foundCellReceb.Row
    End If
    End If
    
    ' Define a planilha onde os dados estão armazenados
    Dim wsCons As Worksheet
    Set wsCons = ThisWorkbook.Worksheets("consConcreto")
    
    ' Definindo a próxima linha disponível na planilha consConcreto
    Dim lastRowCons As Long
    lastRowCons = wsCons.Cells(wsCons.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Escrevendo os dados na planilha na ordem especificada
    wsCons.Cells(lastRowCons, "B").Value = Application.WorksheetFunction.Max(wsCons.Range("B:B")) + 1
    wsCons.Cells(lastRowCons, "C").Value = ListBox1.Value
    wsCons.Cells(lastRowCons, "D").Value = ComboBox1.Value
    wsCons.Cells(lastRowCons, "E").Value = ComboBox2.Value
    wsCons.Cells(lastRowCons, "F").Value = TextBox26.Value
    wsCons.Cells(lastRowCons, "G").Value = Format(TextBox17.Value, "hh:mm:ss")
    wsCons.Cells(lastRowCons, "H").Value = ComboBox4.Value
    wsCons.Cells(lastRowCons, "I").Value = Format(TextBox16.Value, "#0.00")
    wsCons.Cells(lastRowCons, "J").Value = Format(TextBox13.Value, "#0.00")
    wsCons.Cells(lastRowCons, "K").Value = Format(TextBox14.Value, "#0.00")
    wsCons.Cells(lastRowCons, "L").Value = Now()
    wsCons.Cells(lastRowCons, "N").Value = ComboBox6.Value
    wsCons.Cells(lastRowCons, "O").Value = usernameGlobal
    wsCons.Cells(lastRowCons, "P").Value = ComboBox5.Value

    ' Verifica se o valor de TextBox13 é superior ao valor de TextBox25
    If CDbl(TextBox13.Value) > CDbl(TextBox16.Value) Then
    ' Verifica se observacao não está vazia
    If observacao <> "" Then
        ' Escreve a observação na coluna M da mesma linha na planilha consConcreto
        wsCons.Cells(lastRowCons, "M").Value = observacao
    End If
    End If

    
    ' Mensagem de sucesso e fechamento do formulário
    MsgBox "Cadastro realizado com sucesso!", vbInformation, "Sucesso"
    Me.Hide
    ConcConsumo.Show ' Reabre o mesmo formulário
    
    Exit Sub
Error_Handler:
    ' Tratamento de erros
    MsgBox "Ocorreu um erro no cadastro, por favor, verifique suas informações.", vbCritical, "Erro"
    Exit Sub
End Sub

Private Sub CommandButton3_Click()
    Me.Hide 'Fecha o formulário atual (ConcrecebConcreto)
    WelcomeScreen.Show 'Abre o formulário WelcomeScreen
End Sub

Private Sub UserForm_Terminate()
    ' Salvar a pasta de trabalho
    ThisWorkbook.Save
    
    ' Fechar o Excel
    Application.Quit
End Sub

Private Sub Label40_Click()
    ' Abrir o site no navegador padrão
    Shell "cmd /c start https://forms.gle/Lcf3bpwyfFuxoPPs6", vbHide
End Sub