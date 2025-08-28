Attribute VB_Name = "余分スペース削除"
Sub AdvancedSpaceCleaner()

    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim originalValue As Variant
    Dim cleanedValue As String
    Dim processedCount As Long
    Dim changedCount As Long
    Dim startTime As Double
    Dim IncludeFormulas As Boolean
    Dim processFullSheet As Boolean
    
    startTime = Timer
    
    ' エラーハンドリング
    On Error GoTo ErrorHandler
    
    ' アクティブなワークシート取得
    Set ws = Application.ActiveSheet
    If ws Is Nothing Then
        MsgBox "アクティブなワークシートがありません。", vbExclamation, "スペース削除マクロ"
        Exit Sub
    End If
    
    ' シートが保護されているかチェック
    If ws.ProtectContents Then
        MsgBox "シート「" & ws.Name & "」が保護されています。保護を解除してから実行してください。", vbExclamation, "スペース削除マクロ"
        Exit Sub
    End If
    
    ' 選択範囲の確認と設定
    Set rng = Application.selection
    If rng Is Nothing Then
        MsgBox "範囲が選択されていません。", vbExclamation, "スペース削除マクロ"
        Exit Sub
    End If
    
    ' 単一の空セルが選択されている場合
    If rng.count = 1 Then
        If IsEmptyOrError(rng.Cells(1, 1)) Then
            ' 何も選択されていない場合は全シートを対象にするか確認
            If MsgBox("選択されているセルが空です。" & vbCrLf & _
                     "使用されている全範囲を対象にしますか？", _
                     vbYesNo + vbQuestion, "範囲選択") = vbYes Then
                Set rng = ws.UsedRange
                processFullSheet = True
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' 数式セルも処理するか確認
    IncludeFormulas = (MsgBox("数式が入力されているセルも処理しますか？" & vbCrLf & _
                             "「いいえ」を選択すると値のみのセルを処理します。", _
                             vbYesNo + vbQuestion, "処理オプション") = vbYes)
    
    ' 処理対象セル数を表示
    Dim targetCells As Long
    targetCells = CountTargetCells(rng, IncludeFormulas)
    
    If targetCells = 0 Then
        MsgBox "処理対象のセルがありません。", vbInformation, "スペース削除マクロ"
        Exit Sub
    End If
    
    ' 最終確認
    Dim confirmMsg As String
    confirmMsg = "ワークシート: " & ws.Parent.Name & " - " & ws.Name & vbCrLf & _
                "処理対象: " & targetCells & " セル" & vbCrLf & _
                "数式セル: " & IIf(IncludeFormulas, "含む", "除外") & vbCrLf & vbCrLf & _
                "以下の処理を実行します：" & vbCrLf & _
                "・ 前後の空白削除" & vbCrLf & _
                "・ 連続する空白を単一空白に変換" & vbCrLf & _
                "・ 全角・半角空白の統一処理" & vbCrLf & vbCrLf & _
                "実行しますか？"
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "スペース削除マクロ") = vbNo Then
        Exit Sub
    End If
    
    ' 画面更新を停止して処理速度向上
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進行状況表示用
    Dim progressStep As Long
    progressStep = Application.Max(1, targetCells \ 100)
    
    ' メイン処理
    For Each cell In rng
        ' 処理対象セルかチェック
        If ShouldProcessCell(cell, IncludeFormulas) Then
            ' セルの値を安全に取得
            originalValue = GetCellValueSafely(cell)
            
            ' 文字列として処理可能かチェック
            If IsStringValue(originalValue) Then
                cleanedValue = CleanSpaces(CStr(originalValue))
                
                ' 値が変更された場合のみ更新
                If CStr(originalValue) <> cleanedValue Then
                    On Error Resume Next
                    cell.value = cleanedValue
                    If Err.Number = 0 Then
                        changedCount = changedCount + 1
                    End If
                    On Error GoTo ErrorHandler
                End If
            End If
            
            processedCount = processedCount + 1
            
            ' 進行状況表示
            If processedCount Mod progressStep = 0 And targetCells > 1000 Then
                Application.StatusBar = "処理中... " & _
                    Format(processedCount / targetCells, "0%") & _
                    " (" & processedCount & "/" & targetCells & ")"
            End If
        End If
    Next cell
    
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    
    ' 処理結果を報告
    Dim resultMsg As String
    Dim processingTime As Double
    processingTime = Timer - startTime
    
    resultMsg = "スペース削除処理が完了しました。" & vbCrLf & vbCrLf & _
               "処理結果:" & vbCrLf & _
               "・ ワークシート: " & ws.Parent.Name & " - " & ws.Name & vbCrLf & _
               "・ 処理セル数: " & Format(processedCount, "#,##0") & vbCrLf & _
               "・ 変更セル数: " & Format(changedCount, "#,##0") & vbCrLf & _
               "・ 処理時間: " & Format(processingTime, "0.00") & "秒"
    
    MsgBox resultMsg, vbInformation, "処理完了"
    
    Exit Sub
    
ErrorHandler:
    ' エラー時の処理
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.description & vbCrLf & _
           "エラー行: " & Erl, vbCritical, "エラー"
End Sub

Private Function GetCellValueSafely(cell As Range) As Variant
    On Error Resume Next
    GetCellValueSafely = cell.value
    If Err.Number <> 0 Then
        GetCellValueSafely = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function IsStringValue(value As Variant) As Boolean
    On Error Resume Next
    
    ' 空値やエラー値をチェック
    If IsEmpty(value) Or IsError(value) Or IsNull(value) Then
        IsStringValue = False
        Exit Function
    End If
    
    ' 文字列またはテキストに変換可能な値かチェック
    Select Case VarType(value)
        Case vbString
            IsStringValue = True
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            ' 数値だが文字列として扱う場合
            IsStringValue = (Len(Trim(CStr(value))) > 0)
        Case vbDate
            ' 日付も文字列として処理可能
            IsStringValue = True
        Case vbBoolean
            ' ブール値も文字列として処理可能
            IsStringValue = True
        Case Else
            IsStringValue = False
    End Select
    
    On Error GoTo 0
End Function

Private Function IsEmptyOrError(cell As Range) As Boolean
    On Error Resume Next
    Dim cellValue As Variant
    cellValue = cell.value
    
    If Err.Number <> 0 Then
        IsEmptyOrError = True
        Err.Clear
    ElseIf IsEmpty(cellValue) Or IsError(cellValue) Or IsNull(cellValue) Then
        IsEmptyOrError = True
    ElseIf VarType(cellValue) = vbString Then
        IsEmptyOrError = (Len(Trim(cellValue)) = 0)
    Else
        IsEmptyOrError = False
    End If
    
    On Error GoTo 0
End Function

Private Function CleanSpaces(inputText As String) As String
    On Error Resume Next
    
    Dim result As String
    result = inputText
    
    ' 入力値が空の場合はそのまま返す
    If Len(result) = 0 Then
        CleanSpaces = result
        Exit Function
    End If
    
    ' 1. 前後の半角スペース削除
    result = Trim(result)
    
    ' 2. 前後の全角スペース削除
    Do While Len(result) > 0 And Left(result, 1) = "　"
        result = Mid(result, 2)
    Loop
    Do While Len(result) > 0 And Right(result, 1) = "　"
        result = Left(result, Len(result) - 1)
    Loop
    
    ' 3. 連続する半角スペースを単一スペースに
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ' 4. 連続する全角スペースを単一全角スペースに
    Do While InStr(result, "　　") > 0
        result = Replace(result, "　　", "　")
    Loop
    
    ' 5. タブ文字の処理
    result = Replace(result, vbTab, " ")
    
    ' 6. 改行文字の処理（改行を削除する場合）
    ' result = Replace(result, vbCrLf, " ")
    ' result = Replace(result, vbCr, " ")
    ' result = Replace(result, vbLf, " ")
    
    CleanSpaces = result
    
    On Error GoTo 0
End Function

Private Function ShouldProcessCell(cell As Range, IncludeFormulas As Boolean) As Boolean
    On Error Resume Next
    
    ' 空白セルやエラーセルは除外
    If IsEmptyOrError(cell) Then
        ShouldProcessCell = False
        Exit Function
    End If
    
    ' 数式セルの処理判定
    If cell.HasFormula And Not IncludeFormulas Then
        ShouldProcessCell = False
        Exit Function
    End If
    
    ' セルの値を取得して文字列として処理可能かチェック
    Dim cellValue As Variant
    cellValue = GetCellValueSafely(cell)
    
    If Not IsStringValue(cellValue) Then
        ShouldProcessCell = False
        Exit Function
    End If
    
    ShouldProcessCell = True
    
    On Error GoTo 0
End Function

Private Function CountTargetCells(rng As Range, IncludeFormulas As Boolean) As Long
    On Error Resume Next
    
    Dim cell As Range
    Dim count As Long
    
    count = 0
    For Each cell In rng
        If ShouldProcessCell(cell, IncludeFormulas) Then
            count = count + 1
        End If
        
        ' 大量データの場合は途中でカウントを制限
        If count > 100000 Then
            count = count + (rng.count - cell.row + rng.row - 1) ' 概算
            Exit For
        End If
    Next cell
    
    CountTargetCells = count
    
    On Error GoTo 0
End Function

