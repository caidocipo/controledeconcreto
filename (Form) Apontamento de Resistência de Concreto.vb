'APontamento de Resistência de Concreto

Private Sub UserForm_Activate()

    ' Verificar se a variável global "usernameGlobal" está vazia
    If usernameGlobal = "" Then
        ' Fechar o formulário atual
        Me.Hide
        ' Redirecionar para o formulário de login
        Login.Show
    End If
    
    ' Adicionar itens ao ComboBox1
    ComboBox1.Clear
    ComboBox1.AddItem "FCK1"
    ComboBox1.AddItem "FCK3"
    ComboBox1.AddItem "FCK7"
    ComboBox1.AddItem "FCK28"
    
'Adicionando valores ao ComboBox2

    Dim ws2 As Worksheet
    Dim lastRow2 As Long
    Dim cell2 As Range
    Dim Combo2 As Collection
        
    Set ws2 = ThisWorkbook.Sheets("consConcreto") ' Substitua "consConcreto" pelo nome da sua planilha
        
    ' Inicializa a coleção para armazenar os valores únicos
    Set Combo2 = New Collection
        
    ' Obtém a última linha preenchida na coluna A
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
        
    ' Percorre os valores da coluna D a partir da linha 2
    For Each cell2 In ws2.Range("D2:D" & lastRow2)
        ' Verifica se o valor já está na coleção antes de adicioná-lo
        On Error Resume Next
        Combo2.Add cell2.Value, CStr(cell2.Value)
        On Error GoTo 0
    Next cell2
        
    ' Preenche a ComboBox2 com os valores únicos
    For Each Item In Combo2
        ComboBox2.AddItem Item
    Next Item
    
'Adicionando valores ao ComboBox3

    Dim ws3 As Worksheet
    Dim lastRow3 As Long
    Dim cell3 As Range
    Dim Combo3 As Collection
        
    Set ws3 = ThisWorkbook.Sheets("consConcreto") ' Substitua "consConcreto" pelo nome da sua planilha
        
    ' Inicializa a coleção para armazenar os valores únicos
    Set Combo3 = New Collection
        
    ' Obtém a última linha preenchida na coluna A
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row
        
    ' Percorre os valores da coluna E a partir da linha 2
    For Each cell3 In ws3.Range("E2:E" & lastRow3)
        ' Verifica se o valor já está na coleção antes de adicioná-lo
        On Error Resume Next
        Combo3.Add cell3.Value, CStr(cell3.Value)
        On Error GoTo 0
    Next cell3
        
    ' Preenche a ComboBox3 com os valores únicos
    For Each Item In Combo3
        ComboBox3.AddItem Item
    Next Item
        
End Sub

Private Sub ComboBox3_Change()
    ' Limpar a ListBox1
    ListBox1.Clear
    
    ' Verificar se os ComboBoxes estão preenchidos
    If Not IsEmpty(ComboBox2.Value) And Not IsEmpty(ComboBox3.Value) Then
        ' Definir a planilha "consConcreto"
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("consConcreto")
        
        ' Obter os valores selecionados nos ComboBoxes
        Dim searchValue1 As String
        Dim searchValue2 As String
        searchValue1 = ComboBox2.Value
        searchValue2 = ComboBox3.Value
        
        ' Definir as colunas de busca
        Dim rngColumnD As Range
        Dim rngColumnE As Range
        Dim rngColumnC As Range
        
        ' Definir a última linha preenchida nas colunas de busca
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        
        ' Definir os intervalos das colunas de busca
        Set rngColumnD = ws.Range("D2:D" & lastRow)
        Set rngColumnE = ws.Range("E2:E" & lastRow)
        Set rngColumnC = ws.Range("C2:C" & lastRow)
        
        ' Definir o intervalo de resultados
        Dim rngResults As Range
        Set rngResults = ws.Range("C2:C" & lastRow)
        
        ' Definir a variável para armazenar os resultados
        Dim resultArray() As Variant
        Dim resultIndex As Long
        resultIndex = 0
        
        ' Verificar cada célula nas colunas de busca e coletar os resultados correspondentes
        Dim i As Long
        For i = 1 To lastRow - 1
            If rngColumnD.Cells(i).Value = searchValue1 And rngColumnE.Cells(i).Value = searchValue2 Then
                resultIndex = resultIndex + 1
                ReDim Preserve resultArray(1 To resultIndex)
                resultArray(resultIndex) = rngColumnC.Cells(i).Value
            End If
        Next i
        
        ' Verificar se foram encontrados resultados correspondentes
        If resultIndex > 0 Then
            ' Inserir os resultados na ListBox1
            ListBox1.List = Application.WorksheetFunction.Transpose(resultArray)
        Else
            MsgBox "Nenhum resultado encontrado para os valores selecionados.", vbInformation
        End If
    End If
End Sub

Private Sub ListBox1_AfterUpdate()
    ' Define a linha selecionada na ListBox1
    Dim selectedRow As Long
    selectedRow = ListBox1.ListIndex + 2 ' ListIndex começa em 0, então adiciona 1
    
    ' Verifica se há uma linha selecionada
    If selectedRow > 0 Then
        ' Define a planilha onde os dados estão armazenados
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("consConcreto")
        
        ' Define o valor dos TextBoxs como o valor da coluna D da linha selecionada
        TextBox3.Value = ws.Cells(selectedRow, "F").Value
        TextBox4.Value = ws.Cells(selectedRow, "H").Value
        
        ' Obtém o valor selecionado na ListBox1
        Dim selectedValue As String
        selectedValue = ListBox1.Value
        
        ' Define a planilha onde os dados estão armazenados para a ListBox2
        Dim ws2 As Worksheet
        Set ws2 = ThisWorkbook.Worksheets("resisConcreto")
        
        ' Limpa a lista atual do ListBox2
        ListBox2.Clear
        
        ' Define o intervalo de dados na coluna C da planilha resisConcreto
        Dim dataRange As Range
        Set dataRange = ws2.Range("C2:C" & ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Row)
        
        ' Percorre os valores da coluna C e adiciona na ListBox2 aqueles que correspondem ao valor selecionado
        Dim cell As Range
        For Each cell In dataRange
            If cell.Value = selectedValue Then
                ListBox2.AddItem ws2.Cells(cell.Row, "D").Value ' Adiciona o valor correspondente da coluna D na ListBox2
            End If
        Next cell
    End If
End Sub

Private Sub ListBox2_AfterUpdate()
    ' Define a linha selecionada na ListBox2
    Dim selectedRow As Long
    selectedRow = ListBox2.ListIndex + 4 ' ListIndex começa em 0, então adiciona 1
    
    ' Verifica se há uma linha selecionada
    If selectedRow > 0 Then
        ' Define a planilha onde os dados estão armazenados
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("resisConcreto")
        
        ' Define o valor do TextBox6 como o valor da coluna E da linha selecionada
        TextBox6.Value = ws.Cells(selectedRow, "E").Value
    Else
        ' Se nenhum valor estiver selecionado, mantém o TextBox6 vazio
        TextBox6.Value = ""
    End If
End Sub


Private Sub CommandButton1_Click()
    ' Verifica se todos os campos foram preenchidos corretamente
    If ListBox1.ListIndex = -1 Or ComboBox1.ListIndex = -1 Or TextBox5.Value = "" Then
        MsgBox "Por favor, preencha todos os campos.", vbExclamation, "Campos obrigatórios"
        Exit Sub
    End If
    
    On Error GoTo Error_Handler 'Ativa a manipulação de erros
    
    'Definindo a planilha e a próxima linha disponível
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("resisConcreto")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row 'Detecta a última linha preenchida
    
    'Escrevendo os dados na planilha na ordem especificada
    ws.Cells(lastRow + 1, "A").Value = Application.WorksheetFunction.Max(ws.Range("A:A")) + 1
    ws.Cells(lastRow + 1, "C").Value = ListBox1.Value
    ws.Cells(lastRow + 1, "D").Value = ComboBox1.Value
    ws.Cells(lastRow + 1, "E").Value = Format(TextBox5.Value, "#0.00")
    ws.Cells(lastRow + 1, "F").Value = IIf(CheckBox1.Value = True, "True", "False")
    ws.Cells(lastRow + 1, "G").Value = Now()
    ws.Cells(lastRow + 1, "H").Value = usernameGlobal
    
    'Mensagem de sucesso e fechamento do formulário
    MsgBox "Cadastro realizado com sucesso!", vbInformation, "Sucesso"
    Me.Hide
    ConcResist.Show ' Abre o formulário WelcomeScreen
    
    Exit Sub 'Sai do tratamento de erros
    
Error_Handler: 'Tratamento de erros
    'Mensagem de erro e cancelamento do cadastro
    MsgBox "Ocorreu um erro no cadastro, por favor, verifique suas informações.", vbCritical, "Erro"
    Exit Sub 'Sai do tratamento de erros
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