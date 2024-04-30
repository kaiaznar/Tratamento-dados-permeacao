' when (----------------) appear mean difeten modules or buttons'
'---------------------'
Sub Transformar()
    buttontrocatxt.Show
End Sub

'---------------------'
Sub organizadados()
    userorganizadados.Show
End Sub

'---------------------'
Sub Botão2_Clique()
    useformgrafico.Show
End Sub

'---------------------'


Private Sub buttonclear_Click()
    Dim planilha, valorplanilha As String
    planilha = "mediafinal"
    valorplanilha = "0"
    Sheets("dados").Cells = Empty
   
    Do Until valorplanilha = "9"
        Sheets(planilha).Range("A2:B100000") = Empty
        Sheets(planilha).Range("F1:F2") = Empty
        Sheets(planilha).Cells(2, 3) = Empty
        Sheets(planilha).Cells(1, 7) = Empty
        Sheets(planilha).Cells(2, 7) = Empty
        Sheets(planilha).Cells(1, 8) = Empty
        valorplanilha = CStr(valorplanilha + 1)
        planilha = CStr("mediafinal" + valorplanilha)
    Loop
    CommandButton1.ForeColor = &H80000012
    buttondata.ForeColor = &H80000012
    buttonmedia.ForeColor = &H80000012
    buttonsalvar.ForeColor = &H80000012
    tboxendarq.Value = ""
    MsgBox ("Limpinho!!")
End Sub

Private Sub buttondata_Click()
    Dim linha, linha2 As Long
    linha = 10
    linha2 = 11
    Do Until Sheets("dados").Cells(linha, 2).Value = "Time"
        linha = linha + 1
        linha2 = linha2 + 1
    Loop
    If Sheets("dados").Cells(linha, 2).Value = "Time" Then
        Dim do1 As Long
        
        Dim datasdexA(1 To 500000) As Double
        Dim datasdexB(1 To 500000) As String
        Dim datasdexE(1 To 500000) As Single
        Dim datasdexC(1 To 500000) As String
        Dim datasdexD(1 To 500000) As Date
        Dim planilha, valorplanilha As String
        
        do1 = 2
        planilha = "mediafinal"
        valorplanilha = "1"
trocadeplanilha:
        Dim inciodados, fimdados As Long
        Sheets(planilha).Cells(1, 6).Value = Sheets("dados").Cells(linha2, 2).Value
        Sheets(planilha).Cells(1, 7).Value = linha2
        inciodados = linha2
        Do Until Sheets("dados").Cells(linha2, 2).Value = Empty
        
            datasdexC(do1) = Sheets("dados").Cells(linha2, 2)
            Dim for1 As Integer
            For for1 = Len(datasdexC(do1)) To 1 Step -1
                If for1 > 19 Then
                    If Mid(datasdexC(do1), for1, 1) = ":" Then
                    datasdexC(do1) = Left(datasdexC(do1), (for1 - 1))
                    End If
                Else
                    Exit For
                End If
            Next for1
            datasdexD(do1) = Format(DateValue(datasdexC(do1)), "mm/dd/yyyy")
            datasdexA(do1) = datasdexD(do1)
            datasdexE(do1) = datasdexA(do1)
            datasdexB(do1) = CStr(datasdexE(do1))
            If datasdexB(do1 - 1) <> Empty Then
                'SEPARAR POR DIAS
                If datasdexB(do1) <> datasdexB(do1 - 1) Then
                    Sheets(planilha).Cells(2, 6).Value = Sheets("dados").Cells((linha2 - 1), 2).Value
                    Sheets(planilha).Cells(2, 7).Value = linha2 - 1
                    fimdados = linha2
                    Sheets(planilha).Cells(2, 3).Value = (fimdados - inciodados)
                    Sheets(planilha).Cells(1, 8).Value = "continua"
                    planilha = CStr("mediafinal" + valorplanilha)
                    Dim valoressensores(1 To 400), do2 As Integer
                    do2 = 2
                    Do Until Sheets("mediafinal").Cells(do2, 1).Value = Empty
                        Sheets(planilha).Cells(do2, 1).Value = Sheets("mediafinal").Cells(do2, 1).Value
                        do2 = do2 + 1
                    Loop
                    do2 = 2
                    valorplanilha = CStr(valorplanilha + 1)
                    do1 = 2
                    GoTo trocadeplanilha
                
                End If
            End If
                    
                
            linha2 = linha2 + 1
            do1 = do1 + 1
        Loop
        Sheets(planilha).Cells(2, 6).Value = Sheets("dados").Cells((linha2 - 1), 2).Value
        Sheets(planilha).Cells(2, 7).Value = linha2 - 1
        fimdados = linha2
        Sheets(planilha).Cells(2, 3).Value = (fimdados - inciodados)
            
    
    End If
    linha = 10
    linha2 = 11
    buttondata.ForeColor = RGB(112, 219, 147)
    Sheets("mediafinal").Activate
    MsgBox ("Datas finalizadas!")
End Sub

Private Sub buttonmedia_Click()
    
    If Sheets("mediafinal").Cells(1, 6).Value = Empty Then
        MsgBox ("Favor apertar o botão DATAS")
        Exit Sub
    End If
    
    'variáveis da lógica de média
    '
    '
    Dim linha, coluna, erro, auxmed, qntcolunas, qntvazios As Integer
    Dim qntlinha, pulalinha, qntdados As Long
    Dim media, soma As Double
    'var.linha
    linha = 2
    'var.coluna
    coluna = 4
    'var.de espaços errados
    erro = 0
    'var.para auxilio na média
    auxmed = 1
    Dim planilha, valorplanilha As String
    planilha = "mediafinal"
    valorplanilha = "0"
    Dim qntsensores As Integer
    qntsensores = 0
    Do Until Sheets(planilha).Cells(linha, 1).Value = Empty
        qntsensores = qntsensores + 1
        linha = linha + 1
    Loop
    
    qntcolunas = 2 * (qntsensores + 1)
  
   
    'PERCORRE COLUNAS
mudadeplanilha:
    Do Until coluna = qntcolunas + 2

       

        'DIZ O INTERVALO DE DADOS
        Dim iniciodados, finaldados As Long
       
        iniciodados = Sheets(planilha).Cells(1, 7).Value
        finaldados = Sheets(planilha).Cells(2, 7).Value
        pulalinha = iniciodados
        'PERCORRE LINHAS
        Do Until pulalinha = finaldados + 1
        
            'VERIFICAR SE É NÚMERO
            If (IsNumeric(Sheets("dados").Cells(pulalinha, coluna).Value)) Then
                If Sheets("dados").Cells(pulalinha, coluna).Value = 0 Then
                    GoTo casozero
                End If

                'SOMA
                soma = soma + Sheets("dados").Cells(pulalinha, coluna)
                pulalinha = pulalinha + 1
            Else
casozero:
                'CASO NÃO SEJA NENHUMA DAS OPÇÕES ANTERIORES
                erro = erro + 1
                pulalinha = pulalinha + 1
                
                
            End If
            
        Loop
        
        'REALIZA MÉDIA E COLA NA ABA DE MÉDIAS
        media = soma / ((finaldados + 1) - (iniciodados + erro))
        Sheets(planilha).Cells((coluna - (auxmed + 1)), 2) = media
colunasdezero:
        coluna = coluna + 2
        auxmed = auxmed + 1
        soma = 0
        media = 0
        linha = 1
        pulalinha = 1
        erro = 0
    Loop
    valorplanilha = CStr(valorplanilha + 1)
    If Sheets(planilha).Cells(1, 8) = "continua" Then
        planilha = CStr("mediafinal" + valorplanilha)
        coluna = 4
        erro = 0
        auxmed = 1
        GoTo mudadeplanilha
    End If
    buttonmedia.ForeColor = RGB(112, 219, 147)
    MsgBox ("Médias finalizadas!")
End Sub

Private Sub buttonproc_Click()

    Dim intChoice As Integer
    Dim strPath As String
    
    'only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    
    'make the file dialog visible to the user
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    'determine what choice the user made
    If intChoice <> 0 Then
        
        'get the file path selected by the user
        strPath = Application.FileDialog( _
            msoFileDialogOpen).SelectedItems(1)
        
        'print the file path to sheet 1
        tboxendarq.Value = strPath
        
    End If
End Sub
    
Private Sub buttonsalvar_Click()
  If Sheets("mediafinal").Cells(1, 6).Value = Empty Then
        MsgBox ("Favor apertar o botão DATAS")
        Exit Sub
    End If
    If Sheets("mediafinal").Cells(3, 2).Value = Empty Then
        MsgBox ("Favor apertar o botão MÉDIAS")
        Exit Sub
    End If
  
  Dim filedearmazenamento, fileabrir, fileendereco, filepasta, pasta, numero As String
    Dim linha, coluna, a, valorsensor(1 To 100), linha2, b, linha2preenchida, linha2procurando, G As Integer
    Dim qntdados As Long
    Dim data1, data2, data3 As Date
    Dim valormedia(1 To 100) As Single
    linha = 2
    linha2 = 4
    linha2preenchida = 5
    linha2procurando = 4
    coluna = 2
    a = 1
    b = 1
    Dim planilha, valorplanilha As String
    planilha = "mediafinal"
    valorplanilha = "0"
    Dim varpulaabrirarq As Integer
    varpulaabrirarq = 0
    
    'CASO DA CAIXA DE TEXTO
    
    '<> DE VAZIA
    If tboxendarq.Value <> "" Then
    
        Application.ScreenUpdating = False
        
fazerdenovo:
        'ABRE ABA DA PLANILHA ATIVA
        Sheets("mediafinal").Activate
        'VER SE TEM OUTRAS PLANILHAS COM DADOS
        Dim continuacao As String
        continuacao = Sheets(planilha).Cells(1, 8).Value
        'ATRIBUIR INTERVALO DE DATAS
        data1 = Sheets(planilha).Cells(1, 6).Value
        data2 = Sheets(planilha).Cells(2, 6).Value
        qntdados = Sheets(planilha).Cells(2, 3).Value
        
        'ATRIBUIR VALORES DOS SENSORES E RESPECTIVAS MÉDIAS
        Do Until Sheets(planilha).Cells(linha, 2).Value = Empty
            valorsensor(a) = Sheets(planilha).Cells(linha, 1).Value
            valormedia(a) = Sheets(planilha).Cells(linha, 2).Value
            a = a + 1
            linha = linha + 1
        Loop
        
        If varpulaabrirarq = 1 Then
            varpulaabrirarq = 0
            GoTo maisdados
        End If
        
        
        
        ' ABRIR ARQUIVO
        fileendereco = tboxendarq.Value
        For G = Len(fileendereco) To 1 Step -1
            If Mid(fileendereco, G, 1) = "\" Then
                filepasta = Left(fileendereco, G)
                filedearmazenamento = Right(fileendereco, Len(fileendereco) - G)
                Exit For
            End If
        Next G
        
        pasta = InStr(fileendereco, "semana")
        
        numero = Mid(fileendereco, (pasta + 7), 2)
        
        
        fileabrir = Dir(filepasta + "\" + "Médias " + numero + ".xlsx")
         
        
        'ABRIR PASTA DE MÉDIAS DA SEMANA
maisdados:
        Workbooks.Open (fileabrir)

        Sheets("media bruta").Activate
        'PROCURA COLUNA PARA ATRIBUIR OS DADOS
        Do Until coluna = 100
            'CASO JÁ POSSUA DADOS
            If Sheets("media bruta").Cells(1, coluna).Value <> 0 Then
                coluna = coluna + 1
            Else
                'CASO SEJA OS PRIMEIROS DADOS DA PASTA, COLA O NOME DOS SENSORES E OS SEUS RESPECTIVOS DADOS
                If coluna = 2 Then
                
                    Do Until linha2 = (linha + 2)
                        Sheets("media bruta").Cells(linha2, 1).Value = valorsensor(b)
                        Sheets("media bruta").Cells(linha2, 2).Value = valormedia(b)
                        linha2 = linha2 + 1
                        b = b + 1
                    Loop
                    
                    Sheets("media bruta").Cells(1, 2).Value = data1
                    Sheets("media bruta").Cells(2, 2).Value = data2
                    Sheets("media bruta").Cells(3, 2).Value = qntdados
                Else
                    Sheets("media bruta").Cells(1, coluna).Value = data1
                    Sheets("media bruta").Cells(2, coluna).Value = data2
                    Sheets("media bruta").Cells(3, coluna).Value = qntdados
                    
                    'ATRIBUI VALORES DE ACORDO COM O RESPECTIVO SENSOR
                    Do Until linha2 = (linha + 2)
                        'APOS PREENCHER DADOS
                        If valorsensor(b) = "" Then
                            Exit Do
                        End If
                        
                        If Sheets("media bruta").Cells(linha2, 1).Value = valorsensor(b) Then
                            Sheets("media bruta").Cells(linha2, coluna).Value = valormedia(b)
                            linha2 = linha2 + 1
                            b = b + 1
                        'CASO O SENSOR NÃO ESTEJA NA MESMA LINHA
                        Else
                            Do Until linha2procurando = 100
                                'CASO ESTEJA EM OUTRA LINHA
                                If Sheets("media bruta").Cells(linha2procurando, 1).Value = valorsensor(b) Then
                                    Sheets("media bruta").Cells(linha2procurando, coluna).Value = valormedia(b)
                                    linha2procurando = 3
                                    b = b + 1
                                    Exit Do
                                Else
                                    linha2procurando = linha2procurando + 1
                                End If
                            Loop
                            If linha2procurando = 100 Then
                                'CASO O SENSOR NÃO EXISTA
                                Do Until Sheets("media bruta").Cells(linha2preenchida, 1).Value = Empty
                                    linha2preenchida = linha2preenchida + 1
                                Loop
                                Sheets("media bruta").Cells(linha2preenchida, 1).Value = valorsensor(b)
                                Sheets("media bruta").Cells(linha2preenchida, coluna).Value = valormedia(b)
                                linha2preenchida = 5
                                linha2 = linha2 + 1
                                linha2procurando = 4
                                b = b + 1
                            End If
                            
                          
                            
                        End If
                        
                    Loop
                    
                  
                End If
                Exit Do
                
            End If
        Loop
        
        'FAZER MÉDIAS PONDERADAS E COLOCAR EM OUTRA ABA
        Dim lines, columns, subcolumns, zero, totaldados(1 To 500), somatotaldados, fim As Integer
        Dim mediaponderada(1 To 600), valormedponderada(1 To 600), somamediaponderada As Single
        lines = 4
        columns = 2
        subcolumns = 2
        zero = 0
        fim = 2
        somatotaldados = 0
        somamediaponderada = 0
        
        
        
        'ARMAZENAS OS VALORES DAS QNT.DADOS E FAZER A SOMA
        Do Until Sheets("media bruta").Cells(3, subcolumns) = Empty
            totaldados(subcolumns) = Sheets("media bruta").Cells(3, subcolumns)
            somatotaldados = somatotaldados + totaldados(subcolumns)
            subcolumns = subcolumns + 1
        Loop
        'FAZER ATÉ ACABAR OS SENSORES
        Do Until Sheets("media bruta").Cells(lines, 1) = Empty
            Sheets("media ponderada").Cells(lines, 1) = Sheets("media bruta").Cells(lines, 1)
            Do Until fim = subcolumns
                If Sheets("media bruta").Cells(lines, fim) = Empty Then
                    zero = zero + totaldados(fim)
                    fim = fim + 1
                Else
                    fim = fim + 1
                End If
            Loop
            
                
            
            Do Until Sheets("media bruta").Cells(3, columns) = Empty
             
                If Sheets("media bruta").Cells(lines, columns) = Empty Then
                    columns = columns + 1
                    
                Else
                    mediaponderada(columns) = (Sheets("media bruta").Cells(lines, columns) * totaldados(columns)) / (somatotaldados - zero)
                    somamediaponderada = somamediaponderada + mediaponderada(columns)
                   columns = columns + 1
                   
                End If
                
            Loop
            valormedponderada(lines) = somamediaponderada
            Sheets("media ponderada").Cells(lines, 2) = valormedponderada(lines)
            'PINTAR SENSORES 900 VALIDOS DE VERMELHO
            If Sheets("media ponderada").Cells(lines, 1) > 900 Then
                If valormedponderada(lines) > 1 Then
                    With Sheets("media ponderada").Cells(lines, 2).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With Sheets("media ponderada").Cells(lines, 1).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = 255
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    
                
                End If
            End If
            
            columns = 2
            lines = lines + 1
            somamediaponderada = 0
            zero = 0
            fim = 2
        Loop
        
        If continuacao = "continua" Then
            varpulaabrirarq = 1
            valorplanilha = CStr(valorplanilha + 1)
            planilha = CStr("mediafinal" + valorplanilha)
            linha = 2
            linha2 = 4
            linha2preenchida = 5
            linha2procurando = 4
            coluna = 2
            a = 1
            b = 1
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            GoTo fazerdenovo
        End If

        
        'ORGANIZAR PLANILHA COM DADOS E DATAS CERTOS
        'Dim var1for As Integer
        'Dim datasdexE(1 To 600) As Single
        'Dim datasdexC(1 To 600) As String
        'Dim datasdexD(1 To 600) As Date
        'Dim datasdexA(1 To 600) As Double
        'Dim datasdex, datasdexB(1 To 600) As String
        'Dim var1do, var2do, var3do, var4do, var5do, var6do As Integer
        'Dim somadadosdacoluna As Long
        'var1do = 2
        'var2do = 2
        'var3do = 2
        'var4do = 4
        'var6do = 1
        'Do Until Sheets("media bruta").Cells(1, var1do) = Empty
            'datasdexC(var1do) = Sheets("media bruta").Cells(1, var1do)
            
            'For var1for = Len(datasdexC(var1do)) To 1 Step -1
                'If var1for > 19 Then
                    'If Mid(datasdexC(var1do), var1for, 1) = ":" Then
                    'datasdexC(var1do) = Left(datasdexC(var1do), (var1for - 1))
                    'End If
                'Else
                    'Exit For
                'End If
            'Next var1for
            'datasdexD(var1do) = Format(DateValue(datasdexC(var1do)), "mm/dd/yyyy")
            'datasdexA(var1do) = datasdexD(var1do)
            'datasdexE(var1do) = datasdexA(var1do)
            'datasdexB(var1do) = CStr(datasdexE(var1do))
           
            'var1do = var1do + 1
        'Loop
          
        'Do Until var2do = var1do
             'If datasdexB(var2do) = datasdexB(var3do) Then
                'If var3do = var2do Then
                    'var3do = var3do + 1
                    'GoTo igual
                'End If
                'somadadosdacoluna = Sheets("media bruta").Cells(3, var2do) + Sheets("media bruta").Cells(3, var3do)
                  
                'Do Until Sheets("media bruta").Cells(var4do, var2do) = Empty
                    'Sheets("media bruta").Cells(var4do, var2do) = (Sheets("media bruta").Cells(3, var2do) * Sheets("media bruta").Cells(var4do, var2do) + Sheets("media bruta").Cells(3, var3do) * Sheets("media bruta").Cells(var4do, var3do)) / somadadosdacoluna
                    'Sheets("media bruta").Cells(var4do, var3do) = Empty
                    'var4do = var4do + 1
                'Loop
                'Sheets("media bruta").Cells(1, var3do) = Empty
                'Sheets("media bruta").Cells(2, var3do) = Empty
                'Sheets("media bruta").Cells(3, var3do) = Empty
                'Sheets("media bruta").Cells(3, var2do) = somadadosdacoluna
                'var5do = var3do
                'Do Until var5do = var1do - 1
                    'datasdexB(var5do) = datasdexB(var5do + 1)
                    'Do Until Sheets("media bruta").Cells(var6do, var5do + 1) = Empty
                        'Sheets("media bruta").Cells(var6do, var5do) = Sheets("media bruta").Cells(var6do, var5do + 1)
                        'var6do = var6do + 1
                    'Loop
                    'var5do = var5do + 1
                    'If var5do = var1do - 1 Then
                        'Dim var7do As Integer
                        'var7do = 1
                        'Do Until var7do = var6do
                            'Sheets("media bruta").Cells(1, var5do).Delete
                            'var7do = var7do + 1
                        'Loop
                    'End If
                    'var6do = 1
                'Loop
                'var1do = var1do - 1
                'var2do = var2do + 1
                'var4do = 4
                'GoTo igual
                
            'End If
            
            'var3do = var3do + 1
igual:
            'somadadosdacoluna = Empty
            'If var3do = var1do Then
                'var2do = var2do + 1
                'var3do = 2
            'End If
         'Loop
        
        
        
        
        Application.ScreenUpdating = True
        
        ActiveWorkbook.Close
        
        Application.DisplayAlerts = True
        Sheets("mediafinal").Activate
        tboxendarq.Value = ""
    
    Else
        
        MsgBox ("Ensirir o valor da pasta semana!!")
    End If
    buttonsalvar.ForeColor = RGB(112, 219, 147)
End Sub

Private Sub CommandButton1_Click()
    
    Dim fileendereco As String
    
    
    fileendereco = tboxendarq.Value
    
    Application.ScreenUpdating = False
    
    If fileendereco <> "" Then
        fileendereco = "TEXT;" + tboxendarq.Value
        Sheets("dados").Activate
    
        'TEXT;P:\Kai Aznar\missoes\andamento\Tratamento de dados\teste e informações\teste\dados de teste\Dados GE-Permeção INSTR 9 7_26_2016 19_15_18.csv
    
    
        With ActiveSheet.QueryTables.Add(Connection:=fileendereco _
         , Destination:=Range("$A$1"))
         .Name = "Dados GE-Permeção INSTR 9 7_26_2016 19_15_18"
         .FieldNames = True
         .RowNumbers = False
         .FillAdjacentFormulas = False
         .PreserveFormatting = True
         .RefreshOnFileOpen = False
         .RefreshStyle = xlInsertDeleteCells
         .SavePassword = False
         .SaveData = True
         .AdjustColumnWidth = True
         .RefreshPeriod = 0
         .TextFilePromptOnRefresh = False
         .TextFilePlatform = 1252
         .TextFileStartRow = 1
         .TextFileParseType = xlDelimited
         .TextFileTextQualifier = xlTextQualifierDoubleQuote
         .TextFileConsecutiveDelimiter = False
         .TextFileTabDelimiter = True
         .TextFileSemicolonDelimiter = False
         .TextFileCommaDelimiter = True
         .TextFileSpaceDelimiter = False
         .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
         .TextFileDecimalSeparator = "."
         .TextFileThousandsSeparator = " "
         .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
        End With
    
    
    Else
        Application.ScreenUpdating = True
        MsgBox (" E o arquivo seu MERDA?!?!?!")
    End If
    
    
    'CASO ESTEJA SEM DADOS
    Dim recolhadedados As Integer
    Dim fileusado2, filenew2, file1new2, file2new2, file3new2, ordem2 As String
    Dim k2 As Long
    recolhadedados = 1
    Do Until Sheets("dados").Cells(recolhadedados, 1).Value = "Scan  Control:"
        recolhadedados = recolhadedados + 1
    Loop
    If Sheets("dados").Cells((recolhadedados + 1), 1).Value = "" Then
    
        'TROCAR NOME DE ARQUIVO NULO
        
        fileusado2 = tboxendarq.Value
        ordem2 = 1
        
        For k2 = Len(fileusado2) To 1 Step -1
            If Mid(fileusado2, k2, 1) = "\" Then
                file1new2 = Left(fileusado2, k2)
                file2new2 = Right(fileusado2, Len(fileusado2) - k2)
                Exit For
            End If
        Next k2
        
         Do Until ordem2 = 300
            filenew2 = ordem2 + "." + "DADOS-VAZIOS.csv"
            file3new2 = Dir(file1new2 + "\" + filenew2)
            If file3new2 = file2new2 Then
                MsgBox ("Arquivo aberto anteriormente")
                GoTo Saida
            Else
                If file3new2 <> "" Then
                    ordem2 = ordem2 + 1
                Else
                    Name fileusado2 As filenew2
                    Exit Do
                End If
            End If
        Loop
    
    MsgBox ("-----VAZIO----- ABRIR OUTRO ARQUIVO!")
    tboxendarq.Value = ""
    Exit Sub
    GoTo Saida
    End If
    
    'VERIFICAR SE É FORMATO CERTO
    Dim verifica As Integer
    verifica = 1
    Do Until Sheets("dados").Cells(verifica, 4).Value = "Stop Action:"
        verifica = verifica + 1
    Loop
    
    If (IsNumeric(Sheets("dados").Cells((verifica + 2), 4).Value)) Then
        If Sheets("dados").Cells((verifica + 2), 4).Value <> 0 Then
            If Sheets("dados").Cells((verifica + 2), 4).Value <> 1 Then
            
            Else
                MsgBox ("Arquivo de Dados NÃO compatível com o programa!!")
                Modclear.Clear
                GoTo Saida
            
            End If
        
        Else
            MsgBox ("Arquivo de Dados NÃO compatível com o programa!!")
            Modclear.Clear
            GoTo Saida
        End If
    Else
        MsgBox ("Arquivo de Dados NÃO compatível com o programa!!")
        Modclear.Clear
        GoTo Saida
        
    End If
   
    
    
    
    'VARIÁVEIS PARA LOCALIZAR NOMES DE SENSORES
    '
    '
    Dim i, linha As Integer
    linha = 2
    i = 2
    
    ' LOCALIZAR E COLAR NOME DOS SENSORES
    '
    Do Until linha = 200
    
        If (IsNumeric(Sheets("dados").Cells(linha, 1).Value)) Then
        
            Sheets("mediafinal").Cells(i, 1).Value = Sheets("dados").Cells(linha, 1).Value
            i = i + 1
            linha = linha + 1
            
        Else
        
            If (IsNumeric(Sheets("dados").Cells((linha - 1), 1).Value)) Then
                Exit Do
        
            Else
        
                linha = linha + 1
            End If
        End If
    
    Loop
    
    'TROCAR NOME DE ARQUIVO
    Dim fileusado, filenew, file1new, file2new, file3new, ordem As String
    Dim k As Long
    fileusado = tboxendarq.Value
    ordem = 1
    
    For k = Len(fileusado) To 1 Step -1
        If Mid(fileusado, k, 1) = "\" Then
            file1new = Left(fileusado, k)
            file2new = Right(fileusado, Len(fileusado) - k)
            Exit For
        End If
    Next k
    
     Do Until ordem = 300
        filenew = ordem + "." + "DADOS-executado.csv"
        file3new = Dir(file1new + "\" + filenew)
        If file3new = file2new Then
            MsgBox ("Arquivo aberto anteriormente")
            GoTo Saida
        Else
            If file3new <> "" Then
                ordem = ordem + 1
                Else
                Name fileusado As filenew
                Exit Do
            End If
        End If
    Loop
            
    Application.ScreenUpdating = True
Saida:
    CommandButton1.ForeColor = RGB(112, 219, 147)
    Sheets("mediafinal").Activate
    MsgBox ("Transformção finalizada!")
    


End Sub

Private Sub tboxendarq_Change()

End Sub

Private Sub UserForm_Click()

End Sub

'---------------------'
Private Sub buttongraficomedias_Click()
     If tboxarquivomedias.Value = "" Then
        MsgBox ("Enserir arquivo de Médias!")
        
    Else
        Dim file, arquivo As String
        Dim for1 As Integer
        file = tboxarquivomedias.Value
        arquivo = InStr(file, "Médias")
        If arquivo = "0" Then
            MsgBox ("Arquivo NÃO compatível!!")
            tboxarquivomedias.Value = Empty
            cboxaberto.Value = Empty
        Else
            
            
            Workbooks.Open (file)
            
    
            'FAZER GRAFICO DE MÉDIAS DA SEMANA
              Dim vetor1do, vetor2do, vetor3do, vetor4do As Integer
              Dim valorqntdados(1 To 600), valorqntdadossoma As LongLong
              Dim correvalordepontos, correcolunasdedados, correlinhasdossensores, correvalordeserie As Integer
              Dim datasdexA(1 To 600), valoresdograficoA(1 To 600) As Double
              Dim nomegrafico, oi, valoresdografico, valoresdograficoB(1 To 600), datasdex, datasdexB(1 To 600) As String
              Dim varcriargrafabas, varcriargrafabastanques As Integer
              Dim nomesensor As String
              varcriargrafabas = 1
              varcriargrafabastanques = 1
              oi = "1"
              valorqntdadossoma = 0
              correvalordeserie = 1
              correlinhasdossensores = 4
              correvalordepontos = 1
              correcolunasdedados = 2
              vetor1do = 2
              vetor2do = 2
              vetor3do = 2
              vetor4do = 2
              Do Until oi = 20
                  Application.DisplayAlerts = False
                  nomegrafico = CStr("GRÁFICO" + oi)
                  On Error Resume Next
                  Sheets(nomegrafico).Delete
                  oi = CStr(oi + 1)
            
              Loop
volta2:
              Resume Volta
            
            
              
Volta:
              Application.DisplayAlerts = True
              oi = "1"
Volta1:
              Sheets("media bruta").Activate
              'CRIAR GRAFICO
              With Sheets("media bruta").Range("A:Z")
                  ActiveSheet.Shapes.AddChart.Select
                  ActiveChart.Parent.Name = "ph"
                  ActiveChart.ChartType = xlLineMarkers
                  ActiveChart.SetSourceData Source:=Range("$A:$Z")
                  'ActiveChart.ChartTitle.Text = "ph"
                  'PEGAR DADOS
                  If Sheets("media bruta").Cells(correlinhasdossensores, 1) = Empty Then
                    MsgBox ("Arquivo VAZIO!")
                    ActiveWorkbook.Close savechanges:=False
                    tboxarquivomedias.Value = Empty
                    cboxaberto.Value = Empty
                    cboxcategorias.Value = Empty
                    cboxtanque.Value = Empty
                    Exit Sub
                  End If
                    Dim NT As Integer
                    NT = 0
                    Dim var1for As Integer
                    Dim datasdexE(1 To 600) As Single
                    Dim datasdexC(1 To 600) As String
                    Dim datasdexD(1 To 600) As Date
                    Dim var1do, var2do As Integer
                    Dim datasdexAux, valoresdograficoAux As String
                    Dim var2for As Integer
                    Dim correvalordeseriepressao As Integer
                    correvalordeseriepressao = 1
                    Dim correvalordeserietemperatura As Integer
                    correvalordeserietemperatura = 1
                    Dim correvalordeserienivel As Integer
                    correvalordeserienivel = 1
                     Dim correvalordeserieco2 As Integer
                    correvalordeserieco2 = 1
                    Dim correvalordeserieph  As Integer
                    correvalordeserieph = 1
                    Dim correvalordeseriet1, correvalordeseriet2, correvalordeseriet3 As Integer
                    correvalordeseriet1 = 1
                    correvalordeseriet2 = 1
                    correvalordeseriet3 = 1
                    
                  
                     If cboxtanque = Empty Then
                        If cboxcategorias = Empty Then
                            '-----------------GRAFICO DE ATIVOS---------------------------
                            Do Until Sheets("media bruta").Cells(correlinhasdossensores, 1) = Empty
                        
                              If Sheets("media bruta").Cells(correlinhasdossensores, 1) > 900 Then
                                  If correvalordeserie > 10 Then
                                      
                                      
                                      Do Until correvalordeserie = 100
                                          On Error Resume Next
                                          ActiveChart.FullSeriesCollection(correvalordeserie).Delete
                                          If Err.Number = 1004 Then GoTo brigeativo
                    
                                      Loop
brigeativo:
                                      Resume criarplanilhaativo
                                      
criarplanilhaativo:
                                      nomegrafico = CStr("GRÁFICO" + oi)
                                      'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                      
                                      ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                      ActiveChart.ChartTitle.Text = "Gráfico Médias Semana"
                                      ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                      oi = CStr(oi + 1)
                                      correvalordeserie = 1
                                      GoTo Volta1
                                      
                                  End If
                                  
                                  
                                  With ActiveChart
                                      .FullSeriesCollection(correvalordeserie).Name = Sheets("media bruta").Cells(correlinhasdossensores, 1).Value
                                  End With
                                  Do Until Sheets("media bruta").Cells(1, correcolunasdedados) = Empty
                                          
                                      If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) < 1000 Then
                                          If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) > 0 Then
                                              valorqntdados(correcolunasdedados) = Sheets("media bruta").Cells(3, correcolunasdedados).Value
                                              'PEGAR VALORES
                                              valoresdograficoA(correcolunasdedados) = Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados)
                                              valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                          Else
                                              valoresdograficoA(correcolunasdedados) = 0
                                              valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                              valorqntdados(correcolunasdedados) = 0
                                          End If
                                      Else
                                          valoresdograficoA(correcolunasdedados) = 10000
                                          valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                          valorqntdados(correcolunasdedados) = 0
                    
                                      End If
                                      
                                    
                                      
                                      
                                      'AJUSTAR AS DATAS PARA STRING
                                      
                                      
                                      datasdexC(correcolunasdedados) = Sheets("media bruta").Cells(1, correcolunasdedados)
                                      
                                      For var1for = Len(datasdexC(correcolunasdedados)) To 1 Step -1
                                          If var1for > 19 Then
                                              If Mid(datasdexC(correcolunasdedados), var1for, 1) = ":" Then
                                              datasdexC(correcolunasdedados) = Left(datasdexC(correcolunasdedados), (var1for - 1))
                                              End If
                                          Else
                                              Exit For
                                          End If
                                      Next var1for
                                      datasdexD(correcolunasdedados) = Format(DateValue(datasdexC(correcolunasdedados)), "dd/mm/yyyy")
                                      datasdexA(correcolunasdedados) = datasdexD(correcolunasdedados)
                                      datasdexE(correcolunasdedados) = datasdexA(correcolunasdedados)
                                      datasdexB(correcolunasdedados) = CStr(datasdexE(correcolunasdedados))
                              
                                      correcolunasdedados = correcolunasdedados + 1
                            
                                  Loop
                                  'ORGANIZAR AS DAS POR ORDEM
                                var1do = 3
                                var2do = 2
                                  Do Until var1do = correcolunasdedados
                                    
                                    If var1do > var2do Then
                                        If datasdexB(var1do) < datasdexB(var2do) Then
                                            datasdexAux = datasdexB(var1do)
                                            valoresdograficoAux = valoresdograficoB(var1do)
                                            datasdexB(var1do) = datasdexB(var2do)
                                            valoresdograficoB(var1do) = valoresdograficoB(var2do)
                                            datasdexB(var2do) = datasdexAux
                                            valoresdograficoB(var2do) = valoresdograficoAux
                                            datasdexAux = Empty
                                            valoresdograficoAux = Empty
                                        End If
                                    End If
                                    
                                    var2do = var2do + 1
                                    If var2do = correcolunasdedados Then
                                        var1do = var1do + 1
                                        var2do = 2
                                    End If
                                    
                                  Loop
                                  
                              
                                     'TROCAR VIRGULA POR PONTO PARA PODER USAR NO GRAFICO
                                  
                                  Do Until vetor3do = correcolunasdedados
                                  
                                      For var2for = Len(valoresdograficoB(vetor3do)) To 1 Step -1
                                          If Mid(valoresdograficoB(vetor3do), var2for, 1) = "," Then
                                              Mid(valoresdograficoB(vetor3do), var2for, 1) = "."
                                              Exit For
                                          End If
                                      Next var2for
                                      vetor3do = vetor3do + 1
                                  Loop
                                  
                                  'ARMAZENAR DADOS E VALORES DE X EM VOTOR
                                  Do Until vetor1do = correcolunasdedados
                                      If valoresdograficoB(vetor1do) = "pi" Then
                                          GoTo pulavetor1ativo
                                      End If
                                      
                                      If valoresdografico = Empty Then
                                          valoresdografico = "={" + valoresdograficoB(vetor1do)
                                          datasdex = "={" + datasdexB(vetor1do)
                                      Else
                                          valoresdografico = valoresdografico + "," + valoresdograficoB(vetor1do)
                                          datasdex = datasdex + "," + datasdexB(vetor1do)
                                      End If
pulavetor1ativo:
                                      vetor1do = vetor1do + 1
                                      If vetor1do = correcolunasdedados Then
                                          valoresdografico = valoresdografico + "}"
                                          datasdex = datasdex + "}"
                                      End If
                                  
                                      
                                  Loop
                               
                                  
                                  With ActiveChart.FullSeriesCollection(correvalordeserie)
                                      .XValues = datasdex
                                      .Values = valoresdografico
                                  End With
                                  With ActiveChart.Axes(xlCategory, xlPrimary)
                                      .CategoryType = xlTimeScale
                                      .TickLabels.NumberFormat = "dd/mm/yy"
                                  End With
                                  
                                  correvalordeserie = correvalordeserie + 1
                                  correlinhasdossensores = correlinhasdossensores + 1
                                  vetor1do = 2
                                  vetor2do = 2
                                  vetor3do = 2
                                  valorqntdadossoma = 0
                                  correcolunasdedados = 2
                                  valoresdografico = ""
                                  datasdex = ""
                                  
                              Else
                                  correlinhasdossensores = correlinhasdossensores + 1
                              End If
                            Loop
                            
                                
                        
                        Else
                            '-------------------------------GRAFICO DE CATEGORIAS-----------------------------
                            
                            With Sheets("media bruta").Range("A:Z")
                                ActiveSheet.Shapes.AddChart.Select
                                ActiveChart.Parent.Name = "pressao"
                                ActiveChart.ChartType = xlLineMarkers
                                ActiveChart.SetSourceData Source:=Range("$A:$Z")
                                'ActiveChart.ChartTitle.Text = "PRESSÂO"
                            End With
                            With Sheets("media bruta").Range("A:Z")
                                ActiveSheet.Shapes.AddChart.Select
                                ActiveChart.Parent.Name = "temperatura"
                                ActiveChart.ChartType = xlLineMarkers
                                ActiveChart.SetSourceData Source:=Range("$A:$Z")
                                'ActiveChart.ChartTitle.Text = "TEMPERATURA"
                            End With
                            With Sheets("media bruta").Range("A:Z")
                                ActiveSheet.Shapes.AddChart.Select
                                ActiveChart.Parent.Name = "nivel"
                                ActiveChart.ChartType = xlLineMarkers
                                ActiveChart.SetSourceData Source:=Range("$A:$Z")
                                'ActiveChart.ChartTitle.Text = "NÍVEL"
                            End With
                            With Sheets("media bruta").Range("A:Z")
                                ActiveSheet.Shapes.AddChart.Select
                                ActiveChart.Parent.Name = "co2"
                                ActiveChart.ChartType = xlLineMarkers
                                ActiveChart.SetSourceData Source:=Range("$A:$Z")
                                'ActiveChart.ChartTitle.Text = "CO2"
                            End With
                            Do Until Sheets("media bruta").Cells(correlinhasdossensores, 1) = Empty
                                    If Sheets("media bruta").Cells(correlinhasdossensores, 1) > 900 Then
                                      'If Sheets("media bruta").Cells(correlinhasdossensores+1, 1) = Empty Then
                                          
                                          
                                          'Do Until correvalordeseriepressao = 100
                                              'On Error GoTo brige
                                              'ActiveChart.FullSeriesCollection(correvalordeserie).Delete
                        
                                          'Loop
brige:
                                          'Resume criarplanilha
                                          
criarplanilha:
                                          'nomegrafico = CStr("GRÁFICO" + oi)
                                          'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                          
                                          'ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                          'ActiveChart.ChartTitle.Text = "Gráfico Médias Semana"
                                          'ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                          'oi = CStr(oi + 1)
                                          'correvalordeserie = 1
                                          'GoTo Volta1
                                          
                                      'End If
                                      
                                      nomesensor = Sheets("media bruta").Cells(correlinhasdossensores, 1)
                                      
                                      '----------------------------------------SENSORES DE PRESSAO
                                      If nomesensor = 901 Or nomesensor = 903 Or nomesensor = 922 Or nomesensor = 923 Or nomesensor = 904 Or nomesensor = 924 Or nomesensor = 925 Or nomesensor = 905 Or nomesensor = 926 Or nomesensor = 927 Or nomesensor = 906 Or nomesensor = 928 Or nomesensor = 929 Or nomesensor = 907 Or nomesensor = 930 Or nomesensor = 931 Or nomesensor = 914 Or nomesensor = 944 Or nomesensor = 945 Then
                                        NT = 0
                                        ActiveSheet.ChartObjects("pressao").Activate
                                        With ActiveChart
                                          .FullSeriesCollection(correvalordeseriepressao).Name = Sheets("media bruta").Cells(correlinhasdossensores, 1).Value
                                          End With
                                          Do Until Sheets("media bruta").Cells(1, correcolunasdedados) = Empty
                                                  
                                              If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) < 1000 Then
                                                  If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) > 0 Then
                                                      valorqntdados(correcolunasdedados) = Sheets("media bruta").Cells(3, correcolunasdedados).Value
                                                      'PEGAR VALORES
                                                      valoresdograficoA(correcolunasdedados) = Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados)
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  Else
                                                      valoresdograficoA(correcolunasdedados) = 0
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                      valorqntdados(correcolunasdedados) = 0
                                                  End If
                                              Else
                                                  valoresdograficoA(correcolunasdedados) = 10000
                                                  valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  valorqntdados(correcolunasdedados) = 0
                            
                                              End If
                                              
                                            
                                              
                                              
                                              'AJUSTAR AS DATAS PARA STRING
                                             
                                              
                                              datasdexC(correcolunasdedados) = Sheets("media bruta").Cells(1, correcolunasdedados)
                                              
                                              For var1for = Len(datasdexC(correcolunasdedados)) To 1 Step -1
                                                  If var1for > 19 Then
                                                      If Mid(datasdexC(correcolunasdedados), var1for, 1) = ":" Then
                                                      datasdexC(correcolunasdedados) = Left(datasdexC(correcolunasdedados), (var1for - 1))
                                                      End If
                                                  Else
                                                      Exit For
                                                  End If
                                              Next var1for
                                              datasdexD(correcolunasdedados) = Format(DateValue(datasdexC(correcolunasdedados)), "dd/mm/yyyy")
                                              datasdexA(correcolunasdedados) = datasdexD(correcolunasdedados)
                                              datasdexE(correcolunasdedados) = datasdexA(correcolunasdedados)
                                              datasdexB(correcolunasdedados) = CStr(datasdexE(correcolunasdedados))
                                      
                                              correcolunasdedados = correcolunasdedados + 1
                                    
                                          Loop
                                          'ORGANIZAR AS DAS POR ORDEM
                                          
                                          var1do = 3
                                          var2do = 2
                                          Do Until var1do = correcolunasdedados
                                            
                                            If var1do > var2do Then
                                                If datasdexB(var1do) < datasdexB(var2do) Then
                                                    datasdexAux = datasdexB(var1do)
                                                    valoresdograficoAux = valoresdograficoB(var1do)
                                                    datasdexB(var1do) = datasdexB(var2do)
                                                    valoresdograficoB(var1do) = valoresdograficoB(var2do)
                                                    datasdexB(var2do) = datasdexAux
                                                    valoresdograficoB(var2do) = valoresdograficoAux
                                                    datasdexAux = Empty
                                                    valoresdograficoAux = Empty
                                                End If
                                            End If
                                            
                                            var2do = var2do + 1
                                            If var2do = correcolunasdedados Then
                                                var1do = var1do + 1
                                                var2do = 2
                                            End If
                                            
                                          Loop
                                          
                                        
                                         
                                             'TROCAR VIRGULA POR PONTO PARA PODER USAR NO GRAFICO
                                          
                                          Do Until vetor3do = correcolunasdedados
                                          
                                              For var2for = Len(valoresdograficoB(vetor3do)) To 1 Step -1
                                                  If Mid(valoresdograficoB(vetor3do), var2for, 1) = "," Then
                                                      Mid(valoresdograficoB(vetor3do), var2for, 1) = "."
                                                      Exit For
                                                  End If
                                              Next var2for
                                              vetor3do = vetor3do + 1
                                          Loop
                                          
                                          'ARMAZENAR DADOS E VALORES DE X EM VOTOR
                                          Do Until vetor1do = correcolunasdedados
                                              If valoresdograficoB(vetor1do) = "pi" Then
                                                  GoTo pulavetor1pressao
                                              End If
                                              
                                              If valoresdografico = Empty Then
                                                  valoresdografico = "={" + valoresdograficoB(vetor1do)
                                                  datasdex = "={" + datasdexB(vetor1do)
                                              Else
                                                  valoresdografico = valoresdografico + "," + valoresdograficoB(vetor1do)
                                                  datasdex = datasdex + "," + datasdexB(vetor1do)
                                              End If
pulavetor1pressao:
                                              vetor1do = vetor1do + 1
                                              If vetor1do = correcolunasdedados Then
                                                  valoresdografico = valoresdografico + "}"
                                                  datasdex = datasdex + "}"
                                              End If
                                          
                                              
                                          Loop
                                       
                                          ActiveSheet.ChartObjects("pressao").Activate
                                          With ActiveChart.FullSeriesCollection(correvalordeseriepressao)
                                              .XValues = datasdex
                                              .Values = valoresdografico
                                          End With
                                          With ActiveChart.Axes(xlCategory, xlPrimary)
                                              .CategoryType = xlTimeScale
                                              .TickLabels.NumberFormat = "dd/mm/yy"
                                          End With
                                            correvalordeseriepressao = correvalordeseriepressao + 1
                                            correlinhasdossensores = correlinhasdossensores + 1
                                            vetor1do = 2
                                            vetor2do = 2
                                            vetor3do = 2
                                            valorqntdadossoma = 0
                                            correcolunasdedados = 2
                                            valoresdografico = ""
                                            datasdex = ""
                                            nomesensor = Sheets("media bruta").Cells(correlinhasdossensores, 1)
                                            If nomesensor = "" Then
                                                nomesensor = 1
                                            End If
                                      End If
                                      
                                      '------------------------------------------------------SENSORES DE TEMP
                                      If nomesensor = 902 Or nomesensor = 908 Or nomesensor = 932 Or nomesensor = 933 Or nomesensor = 909 Or nomesensor = 934 Or nomesensor = 935 Or nomesensor = 910 Or nomesensor = 936 Or nomesensor = 937 Or nomesensor = 911 Or nomesensor = 938 Or nomesensor = 939 Then
                                        NT = 0
                                        ActiveSheet.ChartObjects("temperatura").Activate
                                        With ActiveChart
                                          .FullSeriesCollection(correvalordeserietemperatura).Name = Sheets("media bruta").Cells(correlinhasdossensores, 1).Value
                                          End With
                                          Do Until Sheets("media bruta").Cells(1, correcolunasdedados) = Empty
                                                  
                                              If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) < 1000 Then
                                                  If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) > 0 Then
                                                      valorqntdados(correcolunasdedados) = Sheets("media bruta").Cells(3, correcolunasdedados).Value
                                                      'PEGAR VALORES
                                                      valoresdograficoA(correcolunasdedados) = Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados)
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  Else
                                                      valoresdograficoA(correcolunasdedados) = 0
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                      valorqntdados(correcolunasdedados) = 0
                                                  End If
                                              Else
                                                  valoresdograficoA(correcolunasdedados) = 10000
                                                  valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  valorqntdados(correcolunasdedados) = 0
                            
                                              End If
                                              
                                            
                                              
                                              
                                              'AJUSTAR AS DATAS PARA STRING
                                             
                                              
                                              datasdexC(correcolunasdedados) = Sheets("media bruta").Cells(1, correcolunasdedados)
                                              
                                              For var1for = Len(datasdexC(correcolunasdedados)) To 1 Step -1
                                                  If var1for > 19 Then
                                                      If Mid(datasdexC(correcolunasdedados), var1for, 1) = ":" Then
                                                      datasdexC(correcolunasdedados) = Left(datasdexC(correcolunasdedados), (var1for - 1))
                                                      End If
                                                  Else
                                                      Exit For
                                                  End If
                                              Next var1for
                                              datasdexD(correcolunasdedados) = Format(DateValue(datasdexC(correcolunasdedados)), "dd/mm/yyyy")
                                              datasdexA(correcolunasdedados) = datasdexD(correcolunasdedados)
                                              datasdexE(correcolunasdedados) = datasdexA(correcolunasdedados)
                                              datasdexB(correcolunasdedados) = CStr(datasdexE(correcolunasdedados))
                                      
                                              correcolunasdedados = correcolunasdedados + 1
                                    
                                          Loop
                                          'ORGANIZAR AS DAS POR ORDEM
                                          
                                          var1do = 3
                                          var2do = 2
                                          Do Until var1do = correcolunasdedados
                                            
                                            If var1do > var2do Then
                                                If datasdexB(var1do) < datasdexB(var2do) Then
                                                    datasdexAux = datasdexB(var1do)
                                                    valoresdograficoAux = valoresdograficoB(var1do)
                                                    datasdexB(var1do) = datasdexB(var2do)
                                                    valoresdograficoB(var1do) = valoresdograficoB(var2do)
                                                    datasdexB(var2do) = datasdexAux
                                                    valoresdograficoB(var2do) = valoresdograficoAux
                                                    datasdexAux = Empty
                                                    valoresdograficoAux = Empty
                                                End If
                                            End If
                                            
                                            var2do = var2do + 1
                                            If var2do = correcolunasdedados Then
                                                var1do = var1do + 1
                                                var2do = 2
                                            End If
                                            
                                          Loop
                                          
                                        
                                         
                                             'TROCAR VIRGULA POR PONTO PARA PODER USAR NO GRAFICO
                                          
                                          Do Until vetor3do = correcolunasdedados
                                          
                                              For var2for = Len(valoresdograficoB(vetor3do)) To 1 Step -1
                                                  If Mid(valoresdograficoB(vetor3do), var2for, 1) = "," Then
                                                      Mid(valoresdograficoB(vetor3do), var2for, 1) = "."
                                                      Exit For
                                                  End If
                                              Next var2for
                                              vetor3do = vetor3do + 1
                                          Loop
                                          
                                          'ARMAZENAR DADOS E VALORES DE X EM VOTOR
                                          Do Until vetor1do = correcolunasdedados
                                              If valoresdograficoB(vetor1do) = "pi" Then
                                                  GoTo pulavetor1temp
                                              End If
                                              
                                              If valoresdografico = Empty Then
                                                  valoresdografico = "={" + valoresdograficoB(vetor1do)
                                                  datasdex = "={" + datasdexB(vetor1do)
                                              Else
                                                  valoresdografico = valoresdografico + "," + valoresdograficoB(vetor1do)
                                                  datasdex = datasdex + "," + datasdexB(vetor1do)
                                              End If
pulavetor1temp:
                                              vetor1do = vetor1do + 1
                                              If vetor1do = correcolunasdedados Then
                                                  valoresdografico = valoresdografico + "}"
                                                  datasdex = datasdex + "}"
                                              End If
                                          
                                              
                                          Loop
                                       
                                          ActiveSheet.ChartObjects("temperatura").Activate
                                          With ActiveChart.FullSeriesCollection(correvalordeserietemperatura)
                                              .XValues = datasdex
                                              .Values = valoresdografico
                                          End With
                                          With ActiveChart.Axes(xlCategory, xlPrimary)
                                              .CategoryType = xlTimeScale
                                              .TickLabels.NumberFormat = "dd/mm/yy"
                                          End With
                                            correvalordeserietemperatura = correvalordeserietemperatura + 1
                                            correlinhasdossensores = correlinhasdossensores + 1
                                            vetor1do = 2
                                            vetor2do = 2
                                            vetor3do = 2
                                            valorqntdadossoma = 0
                                            correcolunasdedados = 2
                                            valoresdografico = ""
                                            datasdex = ""
                                      End If
                                      
                                      '-----------------------------------------SENSORES DE NIVEL
                                      If nomesensor = 913 Or nomesensor = 942 Or nomesensor = 943 Then
                                        NT = 0
                                        ActiveSheet.ChartObjects("nivel").Activate
                                        With ActiveChart
                                          .FullSeriesCollection(correvalordeserienivel).Name = Sheets("media bruta").Cells(correlinhasdossensores, 1).Value
                                          End With
                                          Do Until Sheets("media bruta").Cells(1, correcolunasdedados) = Empty
                                                  
                                              If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) < 1000 Then
                                                  If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) > 0 Then
                                                      valorqntdados(correcolunasdedados) = Sheets("media bruta").Cells(3, correcolunasdedados).Value
                                                      'PEGAR VALORES
                                                      valoresdograficoA(correcolunasdedados) = Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados)
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  Else
                                                      valoresdograficoA(correcolunasdedados) = 0
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                      valorqntdados(correcolunasdedados) = 0
                                                  End If
                                              Else
                                                  valoresdograficoA(correcolunasdedados) = 10000
                                                  valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  valorqntdados(correcolunasdedados) = 0
                            
                                              End If
                                              
                                            
                                              
                                              
                                              'AJUSTAR AS DATAS PARA STRING
                                             
                                              
                                              datasdexC(correcolunasdedados) = Sheets("media bruta").Cells(1, correcolunasdedados)
                                              
                                              For var1for = Len(datasdexC(correcolunasdedados)) To 1 Step -1
                                                  If var1for > 19 Then
                                                      If Mid(datasdexC(correcolunasdedados), var1for, 1) = ":" Then
                                                      datasdexC(correcolunasdedados) = Left(datasdexC(correcolunasdedados), (var1for - 1))
                                                      End If
                                                  Else
                                                      Exit For
                                                  End If
                                              Next var1for
                                              datasdexD(correcolunasdedados) = Format(DateValue(datasdexC(correcolunasdedados)), "dd/mm/yyyy")
                                              datasdexA(correcolunasdedados) = datasdexD(correcolunasdedados)
                                              datasdexE(correcolunasdedados) = datasdexA(correcolunasdedados)
                                              datasdexB(correcolunasdedados) = CStr(datasdexE(correcolunasdedados))
                                      
                                              correcolunasdedados = correcolunasdedados + 1
                                    
                                          Loop
                                          'ORGANIZAR AS DAS POR ORDEM
                                          
                                          var1do = 3
                                          var2do = 2
                                          Do Until var1do = correcolunasdedados
                                            
                                            If var1do > var2do Then
                                                If datasdexB(var1do) < datasdexB(var2do) Then
                                                    datasdexAux = datasdexB(var1do)
                                                    valoresdograficoAux = valoresdograficoB(var1do)
                                                    datasdexB(var1do) = datasdexB(var2do)
                                                    valoresdograficoB(var1do) = valoresdograficoB(var2do)
                                                    datasdexB(var2do) = datasdexAux
                                                    valoresdograficoB(var2do) = valoresdograficoAux
                                                    datasdexAux = Empty
                                                    valoresdograficoAux = Empty
                                                End If
                                            End If
                                            
                                            var2do = var2do + 1
                                            If var2do = correcolunasdedados Then
                                                var1do = var1do + 1
                                                var2do = 2
                                            End If
                                            
                                          Loop
                                          
                                        
                                         
                                             'TROCAR VIRGULA POR PONTO PARA PODER USAR NO GRAFICO
                                          
                                          Do Until vetor3do = correcolunasdedados
                                          
                                              For var2for = Len(valoresdograficoB(vetor3do)) To 1 Step -1
                                                  If Mid(valoresdograficoB(vetor3do), var2for, 1) = "," Then
                                                      Mid(valoresdograficoB(vetor3do), var2for, 1) = "."
                                                      Exit For
                                                  End If
                                              Next var2for
                                              vetor3do = vetor3do + 1
                                          Loop
                                          
                                          'ARMAZENAR DADOS E VALORES DE X EM VOTOR
                                          Do Until vetor1do = correcolunasdedados
                                              If valoresdograficoB(vetor1do) = "pi" Then
                                                  GoTo pulavetor1nivel
                                              End If
                                              
                                              If valoresdografico = Empty Then
                                                  valoresdografico = "={" + valoresdograficoB(vetor1do)
                                                  datasdex = "={" + datasdexB(vetor1do)
                                              Else
                                                  valoresdografico = valoresdografico + "," + valoresdograficoB(vetor1do)
                                                  datasdex = datasdex + "," + datasdexB(vetor1do)
                                              End If
pulavetor1nivel:
                                              vetor1do = vetor1do + 1
                                              If vetor1do = correcolunasdedados Then
                                                  valoresdografico = valoresdografico + "}"
                                                  datasdex = datasdex + "}"
                                              End If
                                          
                                              
                                          Loop
                                       
                                          ActiveSheet.ChartObjects("nivel").Activate
                                          With ActiveChart.FullSeriesCollection(correvalordeserienivel)
                                              .XValues = datasdex
                                              .Values = valoresdografico
                                          End With
                                          With ActiveChart.Axes(xlCategory, xlPrimary)
                                              .CategoryType = xlTimeScale
                                              .TickLabels.NumberFormat = "dd/mm/yy"
                                          End With
                                            correvalordeserienivel = correvalordeserienivel + 1
                                            correlinhasdossensores = correlinhasdossensores + 1
                                            vetor1do = 2
                                            vetor2do = 2
                                            vetor3do = 2
                                            valorqntdadossoma = 0
                                            correcolunasdedados = 2
                                            valoresdografico = ""
                                            datasdex = ""
                                      End If
                                      
                                      '----------------------------------------SENSORES DE CO2
                                      If nomesensor = 917 Then
                                        NT = 0
                                        ActiveSheet.ChartObjects("co2").Activate
                                        With ActiveChart
                                          .FullSeriesCollection(correvalordeserieco2).Name = Sheets("media bruta").Cells(correlinhasdossensores, 1).Value
                                          End With
                                          Do Until Sheets("media bruta").Cells(1, correcolunasdedados) = Empty
                                                  
                                              If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) < 1000 Then
                                                  If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) > 0 Then
                                                      valorqntdados(correcolunasdedados) = Sheets("media bruta").Cells(3, correcolunasdedados).Value
                                                      'PEGAR VALORES
                                                      valoresdograficoA(correcolunasdedados) = Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados)
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  Else
                                                      valoresdograficoA(correcolunasdedados) = 0
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                      valorqntdados(correcolunasdedados) = 0
                                                  End If
                                              Else
                                                  valoresdograficoA(correcolunasdedados) = 10000
                                                  valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  valorqntdados(correcolunasdedados) = 0
                            
                                              End If
                                              
                                            
                                              
                                              
                                              'AJUSTAR AS DATAS PARA STRING
                                             
                                              
                                              datasdexC(correcolunasdedados) = Sheets("media bruta").Cells(1, correcolunasdedados)
                                              
                                              For var1for = Len(datasdexC(correcolunasdedados)) To 1 Step -1
                                                  If var1for > 19 Then
                                                      If Mid(datasdexC(correcolunasdedados), var1for, 1) = ":" Then
                                                      datasdexC(correcolunasdedados) = Left(datasdexC(correcolunasdedados), (var1for - 1))
                                                      End If
                                                  Else
                                                      Exit For
                                                  End If
                                              Next var1for
                                              datasdexD(correcolunasdedados) = Format(DateValue(datasdexC(correcolunasdedados)), "dd/mm/yyyy")
                                              datasdexA(correcolunasdedados) = datasdexD(correcolunasdedados)
                                              datasdexE(correcolunasdedados) = datasdexA(correcolunasdedados)
                                              datasdexB(correcolunasdedados) = CStr(datasdexE(correcolunasdedados))
                                      
                                              correcolunasdedados = correcolunasdedados + 1
                                    
                                          Loop
                                          'ORGANIZAR AS DAS POR ORDEM
                                          
                                          var1do = 3
                                          var2do = 2
                                          Do Until var1do = correcolunasdedados
                                            
                                            If var1do > var2do Then
                                                If datasdexB(var1do) < datasdexB(var2do) Then
                                                    datasdexAux = datasdexB(var1do)
                                                    valoresdograficoAux = valoresdograficoB(var1do)
                                                    datasdexB(var1do) = datasdexB(var2do)
                                                    valoresdograficoB(var1do) = valoresdograficoB(var2do)
                                                    datasdexB(var2do) = datasdexAux
                                                    valoresdograficoB(var2do) = valoresdograficoAux
                                                    datasdexAux = Empty
                                                    valoresdograficoAux = Empty
                                                End If
                                            End If
                                            
                                            var2do = var2do + 1
                                            If var2do = correcolunasdedados Then
                                                var1do = var1do + 1
                                                var2do = 2
                                            End If
                                            
                                          Loop
                                          
                                        
                                         
                                             'TROCAR VIRGULA POR PONTO PARA PODER USAR NO GRAFICO
                                          
                                          Do Until vetor3do = correcolunasdedados
                                          
                                              For var2for = Len(valoresdograficoB(vetor3do)) To 1 Step -1
                                                  If Mid(valoresdograficoB(vetor3do), var2for, 1) = "," Then
                                                      Mid(valoresdograficoB(vetor3do), var2for, 1) = "."
                                                      Exit For
                                                  End If
                                              Next var2for
                                              vetor3do = vetor3do + 1
                                          Loop
                                          
                                          'ARMAZENAR DADOS E VALORES DE X EM VOTOR
                                          Do Until vetor1do = correcolunasdedados
                                              If valoresdograficoB(vetor1do) = "pi" Then
                                                  GoTo pulavetor1co2
                                              End If
                                              
                                              If valoresdografico = Empty Then
                                                  valoresdografico = "={" + valoresdograficoB(vetor1do)
                                                  datasdex = "={" + datasdexB(vetor1do)
                                              Else
                                                  valoresdografico = valoresdografico + "," + valoresdograficoB(vetor1do)
                                                  datasdex = datasdex + "," + datasdexB(vetor1do)
                                              End If
pulavetor1co2:
                                              vetor1do = vetor1do + 1
                                              If vetor1do = correcolunasdedados Then
                                                  valoresdografico = valoresdografico + "}"
                                                  datasdex = datasdex + "}"
                                              End If
                                          
                                              
                                          Loop
                                       
                                          ActiveSheet.ChartObjects("co2").Activate
                                          With ActiveChart.FullSeriesCollection(correvalordeserieco2)
                                              .XValues = datasdex
                                              .Values = valoresdografico
                                          End With
                                          With ActiveChart.Axes(xlCategory, xlPrimary)
                                              .CategoryType = xlTimeScale
                                              .TickLabels.NumberFormat = "dd/mm/yy"
                                          End With
                                            correvalordeserieco2 = correvalordeserieco2 + 1
                                            correlinhasdossensores = correlinhasdossensores + 1
                                            vetor1do = 2
                                            vetor2do = 2
                                            vetor3do = 2
                                            valorqntdadossoma = 0
                                            correcolunasdedados = 2
                                            valoresdografico = ""
                                            datasdex = ""
                                      End If
                                      
                                      '----------------------------------------SENSORES DE PH
                                      If nomesensor = 918 Then
                                        NT = 0
                                        ActiveSheet.ChartObjects("ph").Activate
                                        With ActiveChart
                                          .FullSeriesCollection(correvalordeserieph).Name = Sheets("media bruta").Cells(correlinhasdossensores, 1).Value
                                          End With
                                          Do Until Sheets("media bruta").Cells(1, correcolunasdedados) = Empty
                                                  
                                              If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) < 1000 Then
                                                  If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) > 0 Then
                                                      valorqntdados(correcolunasdedados) = Sheets("media bruta").Cells(3, correcolunasdedados).Value
                                                      'PEGAR VALORES
                                                      valoresdograficoA(correcolunasdedados) = Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados)
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  Else
                                                      valoresdograficoA(correcolunasdedados) = 0
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                      valorqntdados(correcolunasdedados) = 0
                                                  End If
                                              Else
                                                  valoresdograficoA(correcolunasdedados) = 10000
                                                  valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  valorqntdados(correcolunasdedados) = 0
                            
                                              End If
                                              
                                            
                                              
                                              
                                              'AJUSTAR AS DATAS PARA STRING
                                             
                                              
                                              datasdexC(correcolunasdedados) = Sheets("media bruta").Cells(1, correcolunasdedados)
                                              
                                              For var1for = Len(datasdexC(correcolunasdedados)) To 1 Step -1
                                                  If var1for > 19 Then
                                                      If Mid(datasdexC(correcolunasdedados), var1for, 1) = ":" Then
                                                      datasdexC(correcolunasdedados) = Left(datasdexC(correcolunasdedados), (var1for - 1))
                                                      End If
                                                  Else
                                                      Exit For
                                                  End If
                                              Next var1for
                                              datasdexD(correcolunasdedados) = Format(DateValue(datasdexC(correcolunasdedados)), "dd/mm/yyyy")
                                              datasdexA(correcolunasdedados) = datasdexD(correcolunasdedados)
                                              datasdexE(correcolunasdedados) = datasdexA(correcolunasdedados)
                                              datasdexB(correcolunasdedados) = CStr(datasdexE(correcolunasdedados))
                                      
                                              correcolunasdedados = correcolunasdedados + 1
                                    
                                          Loop
                                          'ORGANIZAR AS DAS POR ORDEM
                                          
                                          var1do = 3
                                          var2do = 2
                                          Do Until var1do = correcolunasdedados
                                            
                                            If var1do > var2do Then
                                                If datasdexB(var1do) < datasdexB(var2do) Then
                                                    datasdexAux = datasdexB(var1do)
                                                    valoresdograficoAux = valoresdograficoB(var1do)
                                                    datasdexB(var1do) = datasdexB(var2do)
                                                    valoresdograficoB(var1do) = valoresdograficoB(var2do)
                                                    datasdexB(var2do) = datasdexAux
                                                    valoresdograficoB(var2do) = valoresdograficoAux
                                                    datasdexAux = Empty
                                                    valoresdograficoAux = Empty
                                                End If
                                            End If
                                            
                                            var2do = var2do + 1
                                            If var2do = correcolunasdedados Then
                                                var1do = var1do + 1
                                                var2do = 2
                                            End If
                                            
                                          Loop
                                          
                                        
                                         
                                             'TROCAR VIRGULA POR PONTO PARA PODER USAR NO GRAFICO
                                          
                                          Do Until vetor3do = correcolunasdedados
                                          
                                              For var2for = Len(valoresdograficoB(vetor3do)) To 1 Step -1
                                                  If Mid(valoresdograficoB(vetor3do), var2for, 1) = "," Then
                                                      Mid(valoresdograficoB(vetor3do), var2for, 1) = "."
                                                      Exit For
                                                  End If
                                              Next var2for
                                              vetor3do = vetor3do + 1
                                          Loop
                                          
                                          'ARMAZENAR DADOS E VALORES DE X EM VOTOR
                                          Do Until vetor1do = correcolunasdedados
                                              If valoresdograficoB(vetor1do) = "pi" Then
                                                  GoTo pulavetor1ph
                                              End If
                                              
                                              If valoresdografico = Empty Then
                                                  valoresdografico = "={" + valoresdograficoB(vetor1do)
                                                  datasdex = "={" + datasdexB(vetor1do)
                                              Else
                                                  valoresdografico = valoresdografico + "," + valoresdograficoB(vetor1do)
                                                  datasdex = datasdex + "," + datasdexB(vetor1do)
                                              End If
pulavetor1ph:
                                              vetor1do = vetor1do + 1
                                              If vetor1do = correcolunasdedados Then
                                                  valoresdografico = valoresdografico + "}"
                                                  datasdex = datasdex + "}"
                                              End If
                                          
                                              
                                          Loop
                                       
                                          ActiveSheet.ChartObjects("ph").Activate
                                          With ActiveChart.FullSeriesCollection(correvalordeserieco2)
                                              .XValues = datasdex
                                              .Values = valoresdografico
                                          End With
                                          With ActiveChart.Axes(xlCategory, xlPrimary)
                                              .CategoryType = xlTimeScale
                                              .TickLabels.NumberFormat = "dd/mm/yy"
                                          End With
                                            correvalordeserieph = correvalordeserieph + 1
                                            correlinhasdossensores = correlinhasdossensores + 1
                                            vetor1do = 2
                                            vetor2do = 2
                                            vetor3do = 2
                                            valorqntdadossoma = 0
                                            correcolunasdedados = 2
                                            valoresdografico = ""
                                            datasdex = ""
                                      End If
                                      
                                  '-----------------CASO NAO SEJA SENSORES QUE QUERO-------------------
                                  NT = NT + 1
                                  If NT = 2 Then
                                    correlinhasdossensores = correlinhasdossensores + 1
                                    NT = 0
                                    
                                  End If
                                  
                                  Else
                                      correlinhasdossensores = correlinhasdossensores + 1
                                  End If
                                  
                                  '------------------------GUARDAR GRAFS EM ABAS------------------------
                                   If Sheets("media bruta").Cells(correlinhasdossensores, 1) = Empty Then
                                        
                                        Do Until varcriargrafabas = 6
                                            '
                                            '
                                            '****PRESSAO*****
                                            If varcriargrafabas = 1 Then
                                                ActiveSheet.ChartObjects("pressao").Activate
                                                  Do Until correvalordeseriepressao = 100
                                                      On Error Resume Next
                                                      
                                                      ActiveChart.FullSeriesCollection(correvalordeseriepressao).Delete
                                                      If Err.Number = 1004 Then GoTo brigepres
                                                      
                                
                                                  Loop
brigepres:
                                                  Resume criarplanilhapres
                                                  
criarplanilhapres:
                                                  nomegrafico = CStr("GRÁFICO" + oi)
                                                  'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                                  
                                                  ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                                  ActiveChart.ChartTitle.Text = "Pressão"
                                                  ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                                  oi = CStr(oi + 1)
                                                  varcriargrafabas = varcriargrafabas + 1
                                                  Sheets("media bruta").Activate
                                            End If
                                            
                                            '
                                            '
                                            '*****TEMP*****
                                            If varcriargrafabas = 2 Then
                                                ActiveSheet.ChartObjects("temperatura").Activate
                                                  Do Until correvalordeserietemperatura = 100
                                                      On Error Resume Next
                                                      
                                                      ActiveChart.FullSeriesCollection(correvalordeserietemperatura).Delete
                                                      If Err.Number = 1004 Then GoTo brigetemp
                                                  Loop
brigetemp:
                                                  
                                                  Resume criarplanilhatemp
                                                  
criarplanilhatemp:
                                                  nomegrafico = CStr("GRÁFICO" + oi)
                                                  'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                                  
                                                  ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                                  ActiveChart.ChartTitle.Text = "Temperatura"
                                                  ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                                  oi = CStr(oi + 1)
                                                  varcriargrafabas = varcriargrafabas + 1
                                                  Sheets("media bruta").Activate
                                            End If
                                            '
                                            '
                                            '*****NIVEL*****
                                            If varcriargrafabas = 3 Then
                                                ActiveSheet.ChartObjects("nivel").Activate
                                                  Do Until correvalordeserienivel = 100
                                                      On Error Resume Next
                                                      
                                                      ActiveChart.FullSeriesCollection(correvalordeserienivel).Delete
                                                      If Err.Number = 1004 Then GoTo brigen
                                
                                                  Loop
brigen:
                                                  
                                                  Resume criarplanilhan
                                                  
criarplanilhan:
                                                  nomegrafico = CStr("GRÁFICO" + oi)
                                                  'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                                  
                                                  ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                                  ActiveChart.ChartTitle.Text = "Nível"
                                                  ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                                  oi = CStr(oi + 1)
                                                  varcriargrafabas = varcriargrafabas + 1
                                                  Sheets("media bruta").Activate
                                            End If
                                            '
                                            '
                                            '*****CO2*****
                                            If varcriargrafabas = 4 Then
                                                ActiveSheet.ChartObjects("co2").Activate
                                                  Do Until correvalordeserieco2 = 100
                                                      On Error Resume Next
                                                      
                                                      ActiveChart.FullSeriesCollection(correvalordeserieco2).Delete
                                                      If Err.Number = 1004 Then GoTo brigeco2
                                
                                                  Loop
brigeco2:
                                                  
                                                  Resume criarplanilhaco2
                                                  
criarplanilhaco2:
                                                  nomegrafico = CStr("GRÁFICO" + oi)
                                                  'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                                  
                                                  ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                                  ActiveChart.ChartTitle.Text = "CO2"
                                                  ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                                  oi = CStr(oi + 1)
                                                  varcriargrafabas = varcriargrafabas + 1
                                                  Sheets("media bruta").Activate
                                            End If
                                            '
                                            '
                                            '*****PH*****
                                            If varcriargrafabas = 5 Then
                                                ActiveSheet.ChartObjects("ph").Activate
                                                  Do Until correvalordeserieph = 100
                                                      On Error Resume Next
                                                      
                                                      ActiveChart.FullSeriesCollection(correvalordeserieph).Delete
                                                      If Err.Number = 1004 Then GoTo brigeph
                                
                                                  Loop
brigeph:
                                                  
                                                  Resume criarplanilhaph
                                                  
criarplanilhaph:
                                                  nomegrafico = CStr("GRÁFICO" + oi)
                                                  'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                                  
                                                  ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                                  ActiveChart.ChartTitle.Text = "ph"
                                                  ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                                  oi = CStr(oi + 1)
                                                  varcriargrafabas = varcriargrafabas + 1
                                                  Sheets("media bruta").Activate
                                            End If
                                        Loop
                                         
                                          
                                          
                                    End If
                            Loop
                        End If
                    Else
                    '
                    '
                    '----------------------GRAFICO DE TANQUES-----------------------
                        With Sheets("media bruta").Range("A:Z")
                                ActiveSheet.Shapes.AddChart.Select
                                ActiveChart.Parent.Name = "t1"
                                ActiveChart.ChartType = xlLineMarkers
                                ActiveChart.SetSourceData Source:=Range("$A:$Z")
                                'ActiveChart.ChartTitle.Text = "PRESSÂO"
                            End With
                            With Sheets("media bruta").Range("A:Z")
                                ActiveSheet.Shapes.AddChart.Select
                                ActiveChart.Parent.Name = "t2"
                                ActiveChart.ChartType = xlLineMarkers
                                ActiveChart.SetSourceData Source:=Range("$A:$Z")
                                'ActiveChart.ChartTitle.Text = "TEMPERATURA"
                            End With
                            With Sheets("media bruta").Range("A:Z")
                                ActiveSheet.Shapes.AddChart.Select
                                ActiveChart.Parent.Name = "t3"
                                ActiveChart.ChartType = xlLineMarkers
                                ActiveChart.SetSourceData Source:=Range("$A:$Z")
                                'ActiveChart.ChartTitle.Text = "NÍVEL"
                            End With
                            
                            Do Until Sheets("media bruta").Cells(correlinhasdossensores, 1) = Empty
                                    If Sheets("media bruta").Cells(correlinhasdossensores, 1) > 900 Then
                                      'If Sheets("media bruta").Cells(correlinhasdossensores+1, 1) = Empty Then
                                          
                                          
                                          'Do Until correvalordeseriepressao = 100
                                              'On Error GoTo brige
                                              'ActiveChart.FullSeriesCollection(correvalordeserie).Delete
                        
                                          'Loop
brigea:
                                          'Resume criarplanilha
                                          
criarplanilhaa:
                                          'nomegrafico = CStr("GRÁFICO" + oi)
                                          'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                          
                                          'ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                          'ActiveChart.ChartTitle.Text = "Gráfico Médias Semana"
                                          'ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                          'oi = CStr(oi + 1)
                                          'correvalordeserie = 1
                                          'GoTo Volta1
                                          
                                      'End If
                                      
                                      nomesensor = Sheets("media bruta").Cells(correlinhasdossensores, 1)
                                      
                                      '----------------------------------------SENSORES DE TANQUE 1
                                      If nomesensor = 903 Or nomesensor = 904 Or nomesensor = 905 Or nomesensor = 906 Or nomesensor = 907 Or nomesensor = 914 Or nomesensor = 908 _
                                      Or nomesensor = 909 Or nomesensor = 910 Or nomesensor = 911 Or nomesensor = 913 Then
                                        NT = 0
                                        ActiveSheet.ChartObjects("t1").Activate
                                        With ActiveChart
                                          .FullSeriesCollection(correvalordeseriet1).Name = Sheets("media bruta").Cells(correlinhasdossensores, 1).Value
                                          End With
                                          Do Until Sheets("media bruta").Cells(1, correcolunasdedados) = Empty
                                                  
                                              If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) < 1000 Then
                                                  If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) > 0 Then
                                                      valorqntdados(correcolunasdedados) = Sheets("media bruta").Cells(3, correcolunasdedados).Value
                                                      'PEGAR VALORES
                                                      valoresdograficoA(correcolunasdedados) = Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados)
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  Else
                                                      valoresdograficoA(correcolunasdedados) = 0
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                      valorqntdados(correcolunasdedados) = 0
                                                  End If
                                              Else
                                                  valoresdograficoA(correcolunasdedados) = 10000
                                                  valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  valorqntdados(correcolunasdedados) = 0
                            
                                              End If
                                              
                                            
                                              
                                              
                                              'AJUSTAR AS DATAS PARA STRING
                                             
                                              
                                              datasdexC(correcolunasdedados) = Sheets("media bruta").Cells(1, correcolunasdedados)
                                              
                                              For var1for = Len(datasdexC(correcolunasdedados)) To 1 Step -1
                                                  If var1for > 19 Then
                                                      If Mid(datasdexC(correcolunasdedados), var1for, 1) = ":" Then
                                                      datasdexC(correcolunasdedados) = Left(datasdexC(correcolunasdedados), (var1for - 1))
                                                      End If
                                                  Else
                                                      Exit For
                                                  End If
                                              Next var1for
                                              datasdexD(correcolunasdedados) = Format(DateValue(datasdexC(correcolunasdedados)), "dd/mm/yyyy")
                                              datasdexA(correcolunasdedados) = datasdexD(correcolunasdedados)
                                              datasdexE(correcolunasdedados) = datasdexA(correcolunasdedados)
                                              datasdexB(correcolunasdedados) = CStr(datasdexE(correcolunasdedados))
                                      
                                              correcolunasdedados = correcolunasdedados + 1
                                    
                                          Loop
                                          'ORGANIZAR AS DAS POR ORDEM
                                          
                                          var1do = 3
                                          var2do = 2
                                          Do Until var1do = correcolunasdedados
                                            
                                            If var1do > var2do Then
                                                If datasdexB(var1do) < datasdexB(var2do) Then
                                                    datasdexAux = datasdexB(var1do)
                                                    valoresdograficoAux = valoresdograficoB(var1do)
                                                    datasdexB(var1do) = datasdexB(var2do)
                                                    valoresdograficoB(var1do) = valoresdograficoB(var2do)
                                                    datasdexB(var2do) = datasdexAux
                                                    valoresdograficoB(var2do) = valoresdograficoAux
                                                    datasdexAux = Empty
                                                    valoresdograficoAux = Empty
                                                End If
                                            End If
                                            
                                            var2do = var2do + 1
                                            If var2do = correcolunasdedados Then
                                                var1do = var1do + 1
                                                var2do = 2
                                            End If
                                            
                                          Loop
                                          
                                        
                                         
                                             'TROCAR VIRGULA POR PONTO PARA PODER USAR NO GRAFICO
                                          
                                          Do Until vetor3do = correcolunasdedados
                                          
                                              For var2for = Len(valoresdograficoB(vetor3do)) To 1 Step -1
                                                  If Mid(valoresdograficoB(vetor3do), var2for, 1) = "," Then
                                                      Mid(valoresdograficoB(vetor3do), var2for, 1) = "."
                                                      Exit For
                                                  End If
                                              Next var2for
                                              vetor3do = vetor3do + 1
                                          Loop
                                          
                                          'ARMAZENAR DADOS E VALORES DE X EM VOTOR
                                          Do Until vetor1do = correcolunasdedados
                                              If valoresdograficoB(vetor1do) = "pi" Then
                                                  GoTo pulavetor1t1
                                              End If
                                              
                                              If valoresdografico = Empty Then
                                                  valoresdografico = "={" + valoresdograficoB(vetor1do)
                                                  datasdex = "={" + datasdexB(vetor1do)
                                              Else
                                                  valoresdografico = valoresdografico + "," + valoresdograficoB(vetor1do)
                                                  datasdex = datasdex + "," + datasdexB(vetor1do)
                                              End If
pulavetor1t1:
                                              vetor1do = vetor1do + 1
                                              If vetor1do = correcolunasdedados Then
                                                  valoresdografico = valoresdografico + "}"
                                                  datasdex = datasdex + "}"
                                              End If
                                          
                                              
                                          Loop
                                       
                                          ActiveSheet.ChartObjects("t1").Activate
                                          With ActiveChart.FullSeriesCollection(correvalordeseriet1)
                                              .XValues = datasdex
                                              .Values = valoresdografico
                                          End With
                                          With ActiveChart.Axes(xlCategory, xlPrimary)
                                              .CategoryType = xlTimeScale
                                              .TickLabels.NumberFormat = "dd/mm/yy"
                                          End With
                                            correvalordeseriet1 = correvalordeseriet1 + 1
                                            correlinhasdossensores = correlinhasdossensores + 1
                                            vetor1do = 2
                                            vetor2do = 2
                                            vetor3do = 2
                                            valorqntdadossoma = 0
                                            correcolunasdedados = 2
                                            valoresdografico = ""
                                            datasdex = ""
                                            nomesensor = Sheets("media bruta").Cells(correlinhasdossensores, 1)
                                            If nomesensor = "" Then
                                                nomesensor = 1
                                            End If
                                      End If
                                      
                                      '------------------------------------------------------SENSORES DE TANQUE 2
                                      If nomesensor = 922 Or nomesensor = 924 Or nomesensor = 926 Or nomesensor = 928 Or nomesensor = 930 Or nomesensor = 944 Or nomesensor = 932 _
                                      Or nomesensor = 934 Or nomesensor = 936 Or nomesensor = 938 Or nomesensor = 942 Then
                                        NT = 0
                                        ActiveSheet.ChartObjects("t2").Activate
                                        With ActiveChart
                                          .FullSeriesCollection(correvalordeseriet2).Name = Sheets("media bruta").Cells(correlinhasdossensores, 1).Value
                                          End With
                                          Do Until Sheets("media bruta").Cells(1, correcolunasdedados) = Empty
                                                  
                                              If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) < 1000 Then
                                                  If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) > 0 Then
                                                      valorqntdados(correcolunasdedados) = Sheets("media bruta").Cells(3, correcolunasdedados).Value
                                                      'PEGAR VALORES
                                                      valoresdograficoA(correcolunasdedados) = Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados)
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  Else
                                                      valoresdograficoA(correcolunasdedados) = 0
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                      valorqntdados(correcolunasdedados) = 0
                                                  End If
                                              Else
                                                  valoresdograficoA(correcolunasdedados) = 10000
                                                  valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  valorqntdados(correcolunasdedados) = 0
                            
                                              End If
                                              
                                            
                                              
                                              
                                              'AJUSTAR AS DATAS PARA STRING
                                             
                                              
                                              datasdexC(correcolunasdedados) = Sheets("media bruta").Cells(1, correcolunasdedados)
                                              
                                              For var1for = Len(datasdexC(correcolunasdedados)) To 1 Step -1
                                                  If var1for > 19 Then
                                                      If Mid(datasdexC(correcolunasdedados), var1for, 1) = ":" Then
                                                      datasdexC(correcolunasdedados) = Left(datasdexC(correcolunasdedados), (var1for - 1))
                                                      End If
                                                  Else
                                                      Exit For
                                                  End If
                                              Next var1for
                                              datasdexD(correcolunasdedados) = Format(DateValue(datasdexC(correcolunasdedados)), "dd/mm/yyyy")
                                              datasdexA(correcolunasdedados) = datasdexD(correcolunasdedados)
                                              datasdexE(correcolunasdedados) = datasdexA(correcolunasdedados)
                                              datasdexB(correcolunasdedados) = CStr(datasdexE(correcolunasdedados))
                                      
                                              correcolunasdedados = correcolunasdedados + 1
                                    
                                          Loop
                                          'ORGANIZAR AS DAS POR ORDEM
                                          
                                          var1do = 3
                                          var2do = 2
                                          Do Until var1do = correcolunasdedados
                                            
                                            If var1do > var2do Then
                                                If datasdexB(var1do) < datasdexB(var2do) Then
                                                    datasdexAux = datasdexB(var1do)
                                                    valoresdograficoAux = valoresdograficoB(var1do)
                                                    datasdexB(var1do) = datasdexB(var2do)
                                                    valoresdograficoB(var1do) = valoresdograficoB(var2do)
                                                    datasdexB(var2do) = datasdexAux
                                                    valoresdograficoB(var2do) = valoresdograficoAux
                                                    datasdexAux = Empty
                                                    valoresdograficoAux = Empty
                                                End If
                                            End If
                                            
                                            var2do = var2do + 1
                                            If var2do = correcolunasdedados Then
                                                var1do = var1do + 1
                                                var2do = 2
                                            End If
                                            
                                          Loop
                                          
                                        
                                         
                                             'TROCAR VIRGULA POR PONTO PARA PODER USAR NO GRAFICO
                                          
                                          Do Until vetor3do = correcolunasdedados
                                          
                                              For var2for = Len(valoresdograficoB(vetor3do)) To 1 Step -1
                                                  If Mid(valoresdograficoB(vetor3do), var2for, 1) = "," Then
                                                      Mid(valoresdograficoB(vetor3do), var2for, 1) = "."
                                                      Exit For
                                                  End If
                                              Next var2for
                                              vetor3do = vetor3do + 1
                                          Loop
                                          
                                          'ARMAZENAR DADOS E VALORES DE X EM VOTOR
                                          Do Until vetor1do = correcolunasdedados
                                              If valoresdograficoB(vetor1do) = "pi" Then
                                                  GoTo pulavetor1t2
                                              End If
                                              
                                              If valoresdografico = Empty Then
                                                  valoresdografico = "={" + valoresdograficoB(vetor1do)
                                                  datasdex = "={" + datasdexB(vetor1do)
                                              Else
                                                  valoresdografico = valoresdografico + "," + valoresdograficoB(vetor1do)
                                                  datasdex = datasdex + "," + datasdexB(vetor1do)
                                              End If
pulavetor1t2:
                                              vetor1do = vetor1do + 1
                                              If vetor1do = correcolunasdedados Then
                                                  valoresdografico = valoresdografico + "}"
                                                  datasdex = datasdex + "}"
                                              End If
                                          
                                              
                                          Loop
                                       
                                          ActiveSheet.ChartObjects("t2").Activate
                                          With ActiveChart.FullSeriesCollection(correvalordeseriet2)
                                              .XValues = datasdex
                                              .Values = valoresdografico
                                          End With
                                          With ActiveChart.Axes(xlCategory, xlPrimary)
                                              .CategoryType = xlTimeScale
                                              .TickLabels.NumberFormat = "dd/mm/yy"
                                          End With
                                            correvalordeseriet2 = correvalordeseriet2 + 1
                                            correlinhasdossensores = correlinhasdossensores + 1
                                            vetor1do = 2
                                            vetor2do = 2
                                            vetor3do = 2
                                            valorqntdadossoma = 0
                                            correcolunasdedados = 2
                                            valoresdografico = ""
                                            datasdex = ""
                                      End If
                                      
                                      '-----------------------------------------SENSORES DE TANQUE 3
                                      If nomesensor = 923 Or nomesensor = 925 Or nomesensor = 927 Or nomesensor = 929 Or nomesensor = 931 Or nomesensor = 945 _
                                      Or nomesensor = 933 Or nomesensor = 935 Or nomesensor = 937 Or nomesensor = 939 Or nomesensor = 943 Then
                                        NT = 0
                                        ActiveSheet.ChartObjects("t3").Activate
                                        With ActiveChart
                                          .FullSeriesCollection(correvalordeseriet3).Name = Sheets("media bruta").Cells(correlinhasdossensores, 1).Value
                                          End With
                                          Do Until Sheets("media bruta").Cells(1, correcolunasdedados) = Empty
                                                  
                                              If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) < 1000 Then
                                                  If Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados) > 0 Then
                                                      valorqntdados(correcolunasdedados) = Sheets("media bruta").Cells(3, correcolunasdedados).Value
                                                      'PEGAR VALORES
                                                      valoresdograficoA(correcolunasdedados) = Sheets("media bruta").Cells(correlinhasdossensores, correcolunasdedados)
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  Else
                                                      valoresdograficoA(correcolunasdedados) = 0
                                                      valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                      valorqntdados(correcolunasdedados) = 0
                                                  End If
                                              Else
                                                  valoresdograficoA(correcolunasdedados) = 10000
                                                  valoresdograficoB(correcolunasdedados) = CStr(valoresdograficoA(correcolunasdedados))
                                                  valorqntdados(correcolunasdedados) = 0
                            
                                              End If
                                              
                                            
                                              
                                              
                                              'AJUSTAR AS DATAS PARA STRING
                                             
                                              
                                              datasdexC(correcolunasdedados) = Sheets("media bruta").Cells(1, correcolunasdedados)
                                              
                                              For var1for = Len(datasdexC(correcolunasdedados)) To 1 Step -1
                                                  If var1for > 19 Then
                                                      If Mid(datasdexC(correcolunasdedados), var1for, 1) = ":" Then
                                                      datasdexC(correcolunasdedados) = Left(datasdexC(correcolunasdedados), (var1for - 1))
                                                      End If
                                                  Else
                                                      Exit For
                                                  End If
                                              Next var1for
                                              datasdexD(correcolunasdedados) = Format(DateValue(datasdexC(correcolunasdedados)), "dd/mm/yyyy")
                                              datasdexA(correcolunasdedados) = datasdexD(correcolunasdedados)
                                              datasdexE(correcolunasdedados) = datasdexA(correcolunasdedados)
                                              datasdexB(correcolunasdedados) = CStr(datasdexE(correcolunasdedados))
                                      
                                              correcolunasdedados = correcolunasdedados + 1
                                    
                                          Loop
                                          'ORGANIZAR AS DAS POR ORDEM
                                          
                                          var1do = 3
                                          var2do = 2
                                          Do Until var1do = correcolunasdedados
                                            
                                            If var1do > var2do Then
                                                If datasdexB(var1do) < datasdexB(var2do) Then
                                                    datasdexAux = datasdexB(var1do)
                                                    valoresdograficoAux = valoresdograficoB(var1do)
                                                    datasdexB(var1do) = datasdexB(var2do)
                                                    valoresdograficoB(var1do) = valoresdograficoB(var2do)
                                                    datasdexB(var2do) = datasdexAux
                                                    valoresdograficoB(var2do) = valoresdograficoAux
                                                    datasdexAux = Empty
                                                    valoresdograficoAux = Empty
                                                End If
                                            End If
                                            
                                            var2do = var2do + 1
                                            If var2do = correcolunasdedados Then
                                                var1do = var1do + 1
                                                var2do = 2
                                            End If
                                            
                                          Loop
                                          
                                        
                                         
                                             'TROCAR VIRGULA POR PONTO PARA PODER USAR NO GRAFICO
                                          
                                          Do Until vetor3do = correcolunasdedados
                                          
                                              For var2for = Len(valoresdograficoB(vetor3do)) To 1 Step -1
                                                  If Mid(valoresdograficoB(vetor3do), var2for, 1) = "," Then
                                                      Mid(valoresdograficoB(vetor3do), var2for, 1) = "."
                                                      Exit For
                                                  End If
                                              Next var2for
                                              vetor3do = vetor3do + 1
                                          Loop
                                          
                                          'ARMAZENAR DADOS E VALORES DE X EM VOTOR
                                          Do Until vetor1do = correcolunasdedados
                                              If valoresdograficoB(vetor1do) = "pi" Then
                                                  GoTo pulavetor1t3
                                              End If
                                              
                                              If valoresdografico = Empty Then
                                                  valoresdografico = "={" + valoresdograficoB(vetor1do)
                                                  datasdex = "={" + datasdexB(vetor1do)
                                              Else
                                                  valoresdografico = valoresdografico + "," + valoresdograficoB(vetor1do)
                                                  datasdex = datasdex + "," + datasdexB(vetor1do)
                                              End If
pulavetor1t3:
                                              vetor1do = vetor1do + 1
                                              If vetor1do = correcolunasdedados Then
                                                  valoresdografico = valoresdografico + "}"
                                                  datasdex = datasdex + "}"
                                              End If
                                          
                                              
                                          Loop
                                       
                                          ActiveSheet.ChartObjects("t3").Activate
                                          With ActiveChart.FullSeriesCollection(correvalordeserienivel)
                                              .XValues = datasdex
                                              .Values = valoresdografico
                                          End With
                                          With ActiveChart.Axes(xlCategory, xlPrimary)
                                              .CategoryType = xlTimeScale
                                              .TickLabels.NumberFormat = "dd/mm/yy"
                                          End With
                                            correvalordeseriet3 = correvalordeseriet3 + 1
                                            correlinhasdossensores = correlinhasdossensores + 1
                                            vetor1do = 2
                                            vetor2do = 2
                                            vetor3do = 2
                                            valorqntdadossoma = 0
                                            correcolunasdedados = 2
                                            valoresdografico = ""
                                            datasdex = ""
                                      End If
                                      
                                      
                                  '-----------------CASO NAO SEJA SENSORES QUE QUERO-------------------
                                  NT = NT + 1
                                  If NT = 2 Then
                                    correlinhasdossensores = correlinhasdossensores + 1
                                    NT = 0
                                    
                                  End If
                                  
                                  Else
                                      correlinhasdossensores = correlinhasdossensores + 1
                                  End If
                                  
                                  '------------------------GUARDAR GRAFS EM ABAS------------------------
                                   If Sheets("media bruta").Cells(correlinhasdossensores, 1) = Empty Then
                                        
                                        Do Until varcriargrafabastanques = 4
                                            '
                                            '
                                            '****TANQUE 1*****
                                            If varcriargrafabastanques = 1 Then
                                                ActiveSheet.ChartObjects("t1").Activate
                                                  Do Until correvalordeseriet1 = 100
                                                      On Error Resume Next
                                                      
                                                      ActiveChart.FullSeriesCollection(correvalordeseriet1).Delete
                                                      If Err.Number = 1004 Then GoTo briget1
                                                      
                                
                                                  Loop
briget1:
                                                  Resume criarplanilhat1
                                                  
criarplanilhat1:
                                                  nomegrafico = CStr("GRÁFICO" + oi)
                                                  'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                                  
                                                  ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                                  ActiveChart.ChartTitle.Text = "TANQUE 1"
                                                  ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                                  oi = CStr(oi + 1)
                                                  varcriargrafabastanques = varcriargrafabastanques + 1
                                                  Sheets("media bruta").Activate
                                            End If
                                            
                                            '
                                            '
                                            '*****TANQUE 2*****
                                            If varcriargrafabastanques = 2 Then
                                                ActiveSheet.ChartObjects("t2").Activate
                                                  Do Until correvalordeseriet2 = 100
                                                      On Error Resume Next
                                                      
                                                      ActiveChart.FullSeriesCollection(correvalordeseriet2).Delete
                                                      If Err.Number = 1004 Then GoTo briget2
                                                  Loop
briget2:
                                                  
                                                  Resume criarplanilhat2
                                                  
criarplanilhat2:
                                                  nomegrafico = CStr("GRÁFICO" + oi)
                                                  'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                                  
                                                  ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                                  ActiveChart.ChartTitle.Text = "Tanque 2"
                                                  ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                                  oi = CStr(oi + 1)
                                                  varcriargrafabastanques = varcriargrafabastanques + 1
                                                  Sheets("media bruta").Activate
                                            End If
                                            '
                                            '
                                            '*****TANQUE 3*****
                                            If varcriargrafabastanques = 3 Then
                                                ActiveSheet.ChartObjects("t3").Activate
                                                  Do Until correvalordeseriet3 = 100
                                                      On Error Resume Next
                                                      
                                                      ActiveChart.FullSeriesCollection(correvalordeseriet3).Delete
                                                      If Err.Number = 1004 Then GoTo briget3
                                
                                                  Loop
briget3:
                                                  
                                                  Resume criarplanilhat3
                                                  
criarplanilhat3:
                                                  nomegrafico = CStr("GRÁFICO" + oi)
                                                  'ABRIR E EXCLUIR UMA ABA SE JA EXISTIR
                                                  
                                                  ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                                  ActiveChart.ChartTitle.Text = "TANQUE 3"
                                                  ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                                                  oi = CStr(oi + 1)
                                                  varcriargrafabastanques = varcriargrafabastanques + 1
                                                  Sheets("media bruta").Activate
                                            End If
                                            ActiveSheet.ChartObjects("ph").Delete
                                            
                                            
                                        Loop
                                         
                                          
                                          
                                    End If
                            Loop
                        End If
                    'End If
                  
                  If varcriargrafabas <> 6 Then
                    If varcriargrafabastanques <> 4 Then
                  
                          Do Until correvalordeserie = 100
                              On Error Resume Next
                              ActiveChart.FullSeriesCollection(correvalordeserie).Delete
                              If Err.Number = 1004 Then GoTo brige1
                              
                    
                          Loop
brige1:
                          Resume criarplanilha1
                          
                                      
criarplanilha1:
                          
                          nomegrafico = CStr("GRÁFICO" + oi)
                          
                          
                          ActiveChart.SetElement (msoElementChartTitleAboveChart)
                          ActiveChart.ChartTitle.Text = "Gráfico Médias Semana"
                          ActiveChart.Location where:=xlLocationAsNewSheet, Name:=nomegrafico
                    End If
                  
                    
                End If
                End With
              If cboxaberto.Value <> Empty Then
                ActiveWorkbook.Save
                MsgBox ("Gráficos finalizados!")
                tboxarquivomedias.Value = ""
                cboxaberto.Value = Empty
                Exit Sub
              End If
              Application.ScreenUpdating = True
              
              ActiveWorkbook.Close
              
              Application.DisplayAlerts = True
              Sheets("mediafinal").Activate
              tboxarquivomedias.Value = ""
        End If
    End If

End Sub

Private Sub buttonprocgraf_Click()
    Dim intChoice As Integer
    Dim strPath As String
    
    'only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    
    'make the file dialog visible to the user
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    'determine what choice the user made
    If intChoice <> 0 Then
        
        'get the file path selected by the user
        strPath = Application.FileDialog( _
            msoFileDialogOpen).SelectedItems(1)
        
        'print the file path to sheet 1
        tboxarquivomedias.Value = strPath
        
    End If
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

'---------------------'
Private Sub buttonorganizardados_Click()
If tboxorganizardados.Value = "" Then
        MsgBox ("Enserir arquivo de Médias!")
        
    Else
        Dim file, arquivo As String
        Dim for1 As Integer
        file = tboxorganizardados.Value
        arquivo = InStr(file, "Médias")
        If arquivo = "0" Then
            MsgBox ("Arquivo NÃO compatível!!")
            tboxorganizardados.Value = Empty
        Else
            
            
            Workbooks.Open (file)
            
            Sheets("media bruta").Activate
            
            'ORGANIZAR PLANILHA COM DADOS E DATAS CERTOS
            Dim var1for As Integer
            Dim datasdexE(1 To 600) As Single
            Dim datasdexC(1 To 600) As String
            Dim datasdexD(1 To 600) As Date
            Dim datasdexA(1 To 600) As Double
            Dim datasdex, datasdexB(1 To 600) As String
            Dim var1do, var2do, var3do, var4do, var5do, var6do As Integer
            Dim somadadosdacoluna As Long
            var1do = 2
            var2do = 2
            var3do = 2
            var4do = 4
            var6do = 1
            Do Until Sheets("media bruta").Cells(1, var1do) = Empty
                datasdexC(var1do) = Sheets("media bruta").Cells(1, var1do)
                
                For var1for = Len(datasdexC(var1do)) To 1 Step -1
                    If var1for > 19 Then
                        If Mid(datasdexC(var1do), var1for, 1) = ":" Then
                        datasdexC(var1do) = Left(datasdexC(var1do), (var1for - 1))
                        End If
                    Else
                        Exit For
                    End If
                Next var1for
                datasdexD(var1do) = Format(DateValue(datasdexC(var1do)), "mm/dd/yyyy")
                datasdexA(var1do) = datasdexD(var1do)
                datasdexE(var1do) = datasdexA(var1do)
                datasdexB(var1do) = CStr(datasdexE(var1do))
               
                var1do = var1do + 1
            Loop
              
            Do Until var2do = var1do
                 If datasdexB(var2do) = datasdexB(var3do) Then
                    If var3do = var2do Then
                        var3do = var3do + 1
                        GoTo igual
                    End If
                    somadadosdacoluna = Sheets("media bruta").Cells(3, var2do) + Sheets("media bruta").Cells(3, var3do)
                      
                    Do Until Sheets("media bruta").Cells(var4do, var2do) = Empty
                        Sheets("media bruta").Cells(var4do, var2do) = (Sheets("media bruta").Cells(3, var2do) * Sheets("media bruta").Cells(var4do, var2do) + Sheets("media bruta").Cells(3, var3do) * Sheets("media bruta").Cells(var4do, var3do)) / somadadosdacoluna
                        Sheets("media bruta").Cells(var4do, var3do) = Empty
                        var4do = var4do + 1
                    Loop
                    Sheets("media bruta").Cells(1, var3do) = Empty
                    Sheets("media bruta").Cells(2, var3do) = Empty
                    Sheets("media bruta").Cells(3, var3do) = Empty
                    Sheets("media bruta").Cells(3, var2do) = somadadosdacoluna
                    var5do = var3do
                    Do Until var5do = var1do - 1
                        datasdexB(var5do) = datasdexB(var5do + 1)
                        Do Until Sheets("media bruta").Cells(var6do, var5do + 1) = Empty
                            Sheets("media bruta").Cells(var6do, var5do) = Sheets("media bruta").Cells(var6do, var5do + 1)
                            var6do = var6do + 1
                        Loop
                        var5do = var5do + 1
                        If var5do = var1do - 1 Then
                            Dim var7do As Integer
                            var7do = 1
                            Do Until var7do = var6do
                                Sheets("media bruta").Cells(1, var5do).Delete
                                var7do = var7do + 1
                            Loop
                        End If
                        var6do = 1
                    Loop
                    var1do = var1do - 1
                    var4do = 4
                    GoTo igual
                    
                End If
                
                var3do = var3do + 1
igual:
                somadadosdacoluna = Empty
                If var3do = var1do Then
                    var2do = var2do + 1
                    var3do = 2
                End If
             Loop
        
    
            
              Application.ScreenUpdating = True
              
              ActiveWorkbook.Close
              
              Application.DisplayAlerts = True
              MsgBox ("Organizado!")
              Sheets("mediafinal").Activate
              tboxorganizardados.Value = ""
        End If
    End If
    
End Sub

Private Sub buttonsearch_Click()
  Dim intChoice As Integer
    Dim strPath As String
    
    'only allow the user to select one file
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    
    'make the file dialog visible to the user
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    
    'determine what choice the user made
    If intChoice <> 0 Then
        
        'get the file path selected by the user
        strPath = Application.FileDialog( _
            msoFileDialogOpen).SelectedItems(1)
        
        'print the file path to sheet 1
        tboxorganizardados.Value = strPath
        
    End If
End Sub

Private Sub UserForm_Click()

End Sub

'---------------------'

