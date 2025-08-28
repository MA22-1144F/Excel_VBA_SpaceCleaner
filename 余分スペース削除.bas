Attribute VB_Name = "�]���X�y�[�X�폜"
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
    
    ' �G���[�n���h�����O
    On Error GoTo ErrorHandler
    
    ' �A�N�e�B�u�ȃ��[�N�V�[�g�擾
    Set ws = Application.ActiveSheet
    If ws Is Nothing Then
        MsgBox "�A�N�e�B�u�ȃ��[�N�V�[�g������܂���B", vbExclamation, "�X�y�[�X�폜�}�N��"
        Exit Sub
    End If
    
    ' �V�[�g���ی삳��Ă��邩�`�F�b�N
    If ws.ProtectContents Then
        MsgBox "�V�[�g�u" & ws.Name & "�v���ی삳��Ă��܂��B�ی���������Ă�����s���Ă��������B", vbExclamation, "�X�y�[�X�폜�}�N��"
        Exit Sub
    End If
    
    ' �I��͈͂̊m�F�Ɛݒ�
    Set rng = Application.selection
    If rng Is Nothing Then
        MsgBox "�͈͂��I������Ă��܂���B", vbExclamation, "�X�y�[�X�폜�}�N��"
        Exit Sub
    End If
    
    ' �P��̋�Z�����I������Ă���ꍇ
    If rng.count = 1 Then
        If IsEmptyOrError(rng.Cells(1, 1)) Then
            ' �����I������Ă��Ȃ��ꍇ�͑S�V�[�g��Ώۂɂ��邩�m�F
            If MsgBox("�I������Ă���Z������ł��B" & vbCrLf & _
                     "�g�p����Ă���S�͈͂�Ώۂɂ��܂����H", _
                     vbYesNo + vbQuestion, "�͈͑I��") = vbYes Then
                Set rng = ws.UsedRange
                processFullSheet = True
            Else
                Exit Sub
            End If
        End If
    End If
    
    ' �����Z�����������邩�m�F
    IncludeFormulas = (MsgBox("���������͂���Ă���Z�����������܂����H" & vbCrLf & _
                             "�u�������v��I������ƒl�݂̂̃Z�����������܂��B", _
                             vbYesNo + vbQuestion, "�����I�v�V����") = vbYes)
    
    ' �����ΏۃZ������\��
    Dim targetCells As Long
    targetCells = CountTargetCells(rng, IncludeFormulas)
    
    If targetCells = 0 Then
        MsgBox "�����Ώۂ̃Z��������܂���B", vbInformation, "�X�y�[�X�폜�}�N��"
        Exit Sub
    End If
    
    ' �ŏI�m�F
    Dim confirmMsg As String
    confirmMsg = "���[�N�V�[�g: " & ws.Parent.Name & " - " & ws.Name & vbCrLf & _
                "�����Ώ�: " & targetCells & " �Z��" & vbCrLf & _
                "�����Z��: " & IIf(IncludeFormulas, "�܂�", "���O") & vbCrLf & vbCrLf & _
                "�ȉ��̏��������s���܂��F" & vbCrLf & _
                "�E �O��̋󔒍폜" & vbCrLf & _
                "�E �A������󔒂�P��󔒂ɕϊ�" & vbCrLf & _
                "�E �S�p�E���p�󔒂̓��ꏈ��" & vbCrLf & vbCrLf & _
                "���s���܂����H"
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "�X�y�[�X�폜�}�N��") = vbNo Then
        Exit Sub
    End If
    
    ' ��ʍX�V���~���ď������x����
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' �i�s�󋵕\���p
    Dim progressStep As Long
    progressStep = Application.Max(1, targetCells \ 100)
    
    ' ���C������
    For Each cell In rng
        ' �����ΏۃZ�����`�F�b�N
        If ShouldProcessCell(cell, IncludeFormulas) Then
            ' �Z���̒l�����S�Ɏ擾
            originalValue = GetCellValueSafely(cell)
            
            ' ������Ƃ��ď����\���`�F�b�N
            If IsStringValue(originalValue) Then
                cleanedValue = CleanSpaces(CStr(originalValue))
                
                ' �l���ύX���ꂽ�ꍇ�̂ݍX�V
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
            
            ' �i�s�󋵕\��
            If processedCount Mod progressStep = 0 And targetCells > 1000 Then
                Application.StatusBar = "������... " & _
                    Format(processedCount / targetCells, "0%") & _
                    " (" & processedCount & "/" & targetCells & ")"
            End If
        End If
    Next cell
    
    ' �ݒ�����ɖ߂�
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    
    ' �������ʂ��
    Dim resultMsg As String
    Dim processingTime As Double
    processingTime = Timer - startTime
    
    resultMsg = "�X�y�[�X�폜�������������܂����B" & vbCrLf & vbCrLf & _
               "��������:" & vbCrLf & _
               "�E ���[�N�V�[�g: " & ws.Parent.Name & " - " & ws.Name & vbCrLf & _
               "�E �����Z����: " & Format(processedCount, "#,##0") & vbCrLf & _
               "�E �ύX�Z����: " & Format(changedCount, "#,##0") & vbCrLf & _
               "�E ��������: " & Format(processingTime, "0.00") & "�b"
    
    MsgBox resultMsg, vbInformation, "��������"
    
    Exit Sub
    
ErrorHandler:
    ' �G���[���̏���
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    
    MsgBox "�G���[���������܂����B" & vbCrLf & _
           "�G���[�ԍ�: " & Err.Number & vbCrLf & _
           "�G���[���e: " & Err.description & vbCrLf & _
           "�G���[�s: " & Erl, vbCritical, "�G���["
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
    
    ' ��l��G���[�l���`�F�b�N
    If IsEmpty(value) Or IsError(value) Or IsNull(value) Then
        IsStringValue = False
        Exit Function
    End If
    
    ' ������܂��̓e�L�X�g�ɕϊ��\�Ȓl���`�F�b�N
    Select Case VarType(value)
        Case vbString
            IsStringValue = True
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            ' ���l����������Ƃ��Ĉ����ꍇ
            IsStringValue = (Len(Trim(CStr(value))) > 0)
        Case vbDate
            ' ���t��������Ƃ��ď����\
            IsStringValue = True
        Case vbBoolean
            ' �u�[���l��������Ƃ��ď����\
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
    
    ' ���͒l����̏ꍇ�͂��̂܂ܕԂ�
    If Len(result) = 0 Then
        CleanSpaces = result
        Exit Function
    End If
    
    ' 1. �O��̔��p�X�y�[�X�폜
    result = Trim(result)
    
    ' 2. �O��̑S�p�X�y�[�X�폜
    Do While Len(result) > 0 And Left(result, 1) = "�@"
        result = Mid(result, 2)
    Loop
    Do While Len(result) > 0 And Right(result, 1) = "�@"
        result = Left(result, Len(result) - 1)
    Loop
    
    ' 3. �A�����锼�p�X�y�[�X��P��X�y�[�X��
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    ' 4. �A������S�p�X�y�[�X��P��S�p�X�y�[�X��
    Do While InStr(result, "�@�@") > 0
        result = Replace(result, "�@�@", "�@")
    Loop
    
    ' 5. �^�u�����̏���
    result = Replace(result, vbTab, " ")
    
    ' 6. ���s�����̏����i���s���폜����ꍇ�j
    ' result = Replace(result, vbCrLf, " ")
    ' result = Replace(result, vbCr, " ")
    ' result = Replace(result, vbLf, " ")
    
    CleanSpaces = result
    
    On Error GoTo 0
End Function

Private Function ShouldProcessCell(cell As Range, IncludeFormulas As Boolean) As Boolean
    On Error Resume Next
    
    ' �󔒃Z����G���[�Z���͏��O
    If IsEmptyOrError(cell) Then
        ShouldProcessCell = False
        Exit Function
    End If
    
    ' �����Z���̏�������
    If cell.HasFormula And Not IncludeFormulas Then
        ShouldProcessCell = False
        Exit Function
    End If
    
    ' �Z���̒l���擾���ĕ�����Ƃ��ď����\���`�F�b�N
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
        
        ' ��ʃf�[�^�̏ꍇ�͓r���ŃJ�E���g�𐧌�
        If count > 100000 Then
            count = count + (rng.count - cell.row + rng.row - 1) ' �T�Z
            Exit For
        End If
    Next cell
    
    CountTargetCells = count
    
    On Error GoTo 0
End Function

