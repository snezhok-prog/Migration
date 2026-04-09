Option Explicit

' ===== НАСТРОЙКИ =====
Public Const OUTPUT_FILENAME As String = "catch_orders_registry.json"
Public Const OUTPUT_FILENAME_STRAY As String = "stray_animals_registry.json"
Public Const OUTPUT_FILENAME_CARDS As String = "animal_cards_registry.json"

' Пример не отфильтровываем по тексту,
'       а просто НЕ доходим до него: данные начинаются с 6-й строки.
Public Const SKIP_EXAMPLE_ROW As Boolean = False

Private Const CHUNK_SIZE As Long = 500         ' сколько строк писать в один JSON-файл (0/<=0 = писать всё в один)
Private Const MAX_BASE64_SIZE As Long = 1      ' лимит на base64 в байтах (0 = без лимита)



' ===========================================================
' ================= УНИВЕРСАЛЬНЫЕ ХЕЛПЕРЫ ===================
' ===========================================================

' Вид подблока из скобок: "Мероприятия над животным (Дегельминтизация)" -> "дегельминтизация"
Private Function GroupKind(ByVal s As String) As String
    Dim t As String: t = Norm(s)
    Dim p1 As Long, p2 As Long
    p1 = InStr(t, "("): p2 = InStrRev(t, ")")
    If p1 > 0 And p2 > p1 Then
        GroupKind = Trim$(Mid$(t, p1 + 1, p2 - p1 - 1))
    Else
        GroupKind = t
    End If
End Function

' Матч по виду (достаточно подстроки вида: "дегельминтизац", "вакцинац", "стерилизац", "нанесение идентификационн")
Private Function IsGroupMatch(ByVal s As String, ByVal want As String) As Boolean
    IsGroupMatch = (InStr(GroupKind(s), Norm(want)) > 0)
End Function


Private Function NewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    d.CompareMode = vbTextCompare
    On Error GoTo 0
    Set NewDict = d
End Function

' Универсальный парсер повторяющихся блоков (якорь + строгий подблок + защита от пустых повторов)
Private Function ParseEventGroupStrict( _
    ByVal ws As Worksheet, ByVal r As Long, ByVal spec As Collection, _
    ByVal groupName As String, ByVal fieldMap As Object, _
    ByVal anchorKey As String, _
    Optional ByVal fileHeader As String = "", _
    Optional ByVal fileJsonBase As String = "" _
) As Collection
    
    Dim res As New Collection
    Dim i As Long, col As Long, h As String, s As String
    Dim cur As Object: Set cur = Nothing
    Dim currentType As String, v As Variant
    
    For i = 1 To spec.Count
        s = spec(i)("sNorm")
        If Not IsGroupMatch(s, groupName) Then GoTo NextI
        
        h = spec(i)("hNorm")
        col = spec(i)("col")
        
        ' Тип мероприятия — задаёт currentType ТОЛЬКО когда непустое значение
        If h = Norm("Тип мероприятия") Then
            v = ws.Cells(r, col).value
            If Not IsEmptyLike(v) Then
                currentType = CStr(v)
                ' Закрываем предыдущий, если там был полезный payload
                If Not cur Is Nothing Then
                    If HasPayload(cur, anchorKey) Then res.Add cur
                End If
                Set cur = NewDict()
                SafePut cur, "type", currentType
            End If
            GoTo MaybeFiles
        End If
        
        If fieldMap.Exists(h) Then
            Dim key As String: key = CStr(fieldMap(h))
            v = ws.Cells(r, col).value
            
            ' Новый повтор ТОЛЬКО на якоре и ТОЛЬКО если значение непустое
            If key = anchorKey And Not IsEmptyLike(v) Then
                If Not cur Is Nothing Then
                    If HasPayload(cur, anchorKey) Then res.Add cur
                End If
                Set cur = NewDict()
                If LenB(currentType) > 0 Then SafePut cur, "type", currentType
            End If
            
            If cur Is Nothing Then
                Set cur = NewDict()
                If LenB(currentType) > 0 Then SafePut cur, "type", currentType
            End If
            
            Select Case True
                Case Right$(key, 4) = "date": AddIfPresentDate ws, r, col, cur, key
                Case key = "time":            AddIfPresentTime ws, r, col, cur, key
                Case Else:                    AddIfPresent ws, r, col, cur, key
            End Select
        End If
        
MaybeFiles:
        If LenB(fileHeader) > 0 Then
            If spec(i)("hNorm") = Norm(fileHeader) Then
                ReadFilesIntoObject ws.Cells(r, col).value, cur, fileJsonBase, fileJsonBase & "s"
            End If
        End If
NextI:
    Next i
    
    If Not cur Is Nothing Then
        If HasPayload(cur, anchorKey) Then res.Add cur
    End If
    
    Set ParseEventGroupStrict = res
End Function


' Нормализация текста для сопоставления (заголовки/подзаголовки)
Private Function Norm(ByVal s As String) As String
    Dim t As String: t = LCase$(Trim$(CStr(s)))

    ' --- NEW: приводим все варианты дефисов/тире к обычному "-" ---
    t = Replace$(t, ChrW(&H2010), "-") ' Hyphen
    t = Replace$(t, ChrW(&H2011), "-") ' Non-breaking hyphen
    t = Replace$(t, ChrW(&H2012), "-") ' Figure dash
    t = Replace$(t, ChrW(&H2013), "-") ' En dash
    t = Replace$(t, ChrW(&H2014), "-") ' Em dash
    t = Replace$(t, ChrW(&H2212), "-") ' Minus sign

    t = Replace$(t, vbCrLf, " ")
    t = Replace$(t, vbCr, " ")
    t = Replace$(t, vbLf, " ")
    t = Replace$(t, Chr$(160), " ")
    t = Replace$(t, """", "")
    Do While InStr(t, "  ") > 0: t = Replace$(t, "  ", " "): Loop
    Norm = t
End Function

' NEW: проверка соответствия подблока по одному или нескольким шаблонам ("a|b|c")
Private Function SubblockMatch(ByVal sNorm As String, ByVal patterns As String) As Boolean
    Dim arr() As String, i As Long, p As String
    arr = Split(CStr(patterns), "|")

    For i = LBound(arr) To UBound(arr)
        p = Norm(arr(i))
        If Len(p) > 0 Then
            If InStr(1, sNorm, p, vbTextCompare) > 0 Then
                SubblockMatch = True
                Exit Function
            End If
        End If
    Next i
End Function



' Пусто ли значение?
Private Function IsEmptyLike(v As Variant) As Boolean
    If IsObject(v) Then
        If v Is Nothing Then
            IsEmptyLike = True
        ElseIf TypeName(v) = "Collection" Or TypeName(v) = "Dictionary" Then
            On Error Resume Next
            IsEmptyLike = (v.Count = 0)
            On Error GoTo 0
        Else
            IsEmptyLike = False
        End If
    Else
        Select Case VarType(v)
            Case vbEmpty, vbNull
                IsEmptyLike = True
            Case vbString
                IsEmptyLike = (Trim$(CStr(v)) = "")
            Case Else
                IsEmptyLike = False
        End Select
    End If
End Function

' Безопасная вставка: пропускает пустое, для Dictionary — Add/перезапись, для Collection — добавляет элемент (без ключа)
Private Sub SafePut(ByVal target As Object, ByVal jsonKey As String, ByVal value As Variant)
    If IsEmptyLike(value) Then Exit Sub
    Select Case TypeName(target)
        Case "Dictionary"
            Dim d As Object: Set d = target
            If d.Exists(jsonKey) Then
                d(jsonKey) = value
            Else
                d.Add jsonKey, value
            End If
        Case "Collection"
            ' В коллекции нет ключей — добавляем как есть
            target.Add value
        Case Else
            ' На всякий случай пытаемся через default свойство Item
            On Error Resume Next
            target(jsonKey) = value
            If Err.Number <> 0 Then
                Err.Clear
                CallByName target, "Add", VbMethod, jsonKey, value
            End If
            On Error GoTo 0
    End Select
End Sub

' Упростители для чтения ячеек
Private Sub AddIfPresent(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long, ByVal target As Object, ByVal jsonKey As String)
    Dim v As Variant
    v = ws.Cells(r, col).value
    If IsEmptyLike(v) Then Exit Sub

    ' Если Excel распознал значение как дату, а поле не "Date/Time" (мы сюда попадаем
    ' только для обычных полей), сохраняем отображаемый текст, а не дату.
    If VarType(v) = vbDate Then
        SafePut target, jsonKey, ws.Cells(r, col).Text
    Else
        SafePut target, jsonKey, v
    End If
End Sub

' [IMPROVE] Для ОГРН/ИНН: берём текст из ячейки и оставляем только цифры.
Private Sub AddIfPresentDigits(ByVal ws As Worksheet, _
                               ByVal r As Long, _
                               ByVal col As Long, _
                               ByVal target As Object, _
                               ByVal jsonKey As String)
    Dim rawText As String
    Dim i As Long
    Dim ch As String
    Dim digits As String
    
    rawText = CStr(ws.Cells(r, col).Text)   ' именно .Text, чтобы не словить экспоненту/округление
    For i = 1 To Len(rawText)
        ch = Mid$(rawText, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        End If
    Next i
    
    If Len(Trim$(digits)) = 0 Then Exit Sub
    ' можно при желании добавить проверку длины (13/15 для ОГРН, 10/12 для ИНН),
    ' но для миграции нам важнее не потерять значение, чем жёстко отфильтровать.
    SafePut target, jsonKey, digits
End Sub



Private Sub AddIfPresentDate(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long, ByVal target As Object, ByVal jsonKey As String)
    Dim v As Variant: v = ws.Cells(r, col).value
    Dim s As String: s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Sub
    If IsDate(v) Then
        SafePut target, jsonKey, Format$(CDate(v), "yyyy-mm-dd")
    ElseIf IsDate(s) Then
        SafePut target, jsonKey, Format$(CDate(s), "yyyy-mm-dd")
    Else
        SafePut target, jsonKey, s
    End If
End Sub

Private Sub AddIfPresentTime(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long, ByVal target As Object, ByVal jsonKey As String)
    Dim v As Variant: v = ws.Cells(r, col).value
    If IsEmptyLike(v) Then Exit Sub
    Dim s As String
    If IsNumeric(v) Then
        s = Format$(CDbl(v), "hh:mm")
    Else
        s = CStr(v)
        s = Replace(s, ",", ".")
        If IsNumeric(s) Then s = Format$(CDbl(s), "hh:mm")
    End If
    SafePut target, jsonKey, s
End Sub


' ---------- Поиск листа / границы данных ----------
Private Function FindSheetLike(ByVal pattern As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LCase$(ws.Name) Like LCase$(pattern) Then
            Set FindSheetLike = ws
            Exit Function
        End If
    Next ws
    Set FindSheetLike = Nothing
End Function

Private Function FindHeaderRow(ws As Worksheet, marker As String) As Long
    Dim f As Range
    Set f = ws.Columns(1).Find(What:=marker, LookAt:=xlWhole, LookIn:=xlValues, _
                            SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If Not f Is Nothing Then FindHeaderRow = f.row Else FindHeaderRow = 0
End Function

Private Function GuessDataStartRow(ws As Worksheet, headerRow As Long) As Long
    ' По шаблону все данные начинаются с 6-й строки.
    GuessDataStartRow = 6
End Function


Private Function FindLastRowAny(ws As Worksheet) As Long
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                 MatchCase:=False)
    If lastCell Is Nothing Then FindLastRowAny = 0 Else FindLastRowAny = lastCell.row
End Function

Private Function IsRowEmpty(ws As Worksheet, r As Long, headerRow As Long) As Boolean
    Dim lastCol As Long: lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    IsRowEmpty = (Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))) = 0)
End Function

' ---------- Файлы / base64 ----------
Private Function SplitMany(ByVal s As String) As Collection
    Dim c As New Collection, t As String
    t = Trim$(Replace(Replace(Replace(CStr(s), vbCrLf, vbLf), vbCr, vbLf), vbTab, " "))
    If Len(t) = 0 Then Set SplitMany = Nothing: Exit Function
    t = Replace(t, "|", ";"): t = Replace(t, ",", ";")
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    t = Replace(t, vbLf, ";")
    Dim parts() As String: parts = Split(t, ";")
    Dim i As Long, token As String
    For i = LBound(parts) To UBound(parts)
        token = Trim$(parts(i))
        If Len(token) > 0 Then c.Add token
    Next i
    If c.Count = 0 Then Set SplitMany = Nothing Else Set SplitMany = c
End Function

Private Function FullPathFromToken(ByVal token As String) As String
    Dim baseDir As String: baseDir = ThisWorkbook.path
    If Len(baseDir) = 0 Then
        FullPathFromToken = token
    Else
        FullPathFromToken = baseDir & "\" & token
    End If
End Function

Private Function DesktopDir() As String
    DesktopDir = Environ$("USERPROFILE") & "\Desktop"
End Function

Private Function CleanFileToken(ByVal token As String) As String
    Dim t As String
    t = Trim$(CStr(token))
    t = Replace$(t, """", "")
    t = Replace$(t, "'", "")
    t = Replace$(t, "/", "\")
    Do While InStr(t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop
    CleanFileToken = t
End Function

Private Function IsAbsolutePath(ByVal p As String) As Boolean
    Dim t As String: t = CleanFileToken(p)
    If Len(t) = 0 Then
        IsAbsolutePath = False
        Exit Function
    End If

    If Left$(t, 2) = "\\" Then IsAbsolutePath = True: Exit Function   ' UNC
    If InStr(1, t, ":\", vbTextCompare) > 0 Then IsAbsolutePath = True: Exit Function ' C:\...
    If Left$(t, 1) = "\" Then IsAbsolutePath = True: Exit Function    ' \folder\file (на текущем диске)

    IsAbsolutePath = False
End Function

' Ищем файл
Private Function ResolveFilePath(ByVal token As String) As String
    Dim t As String: t = CleanFileToken(token)
    If Len(t) = 0 Then ResolveFilePath = "": Exit Function

    ' 1) как есть
    If FileExists(t) Then ResolveFilePath = t: Exit Function

    ' если токен абсолютный, но не найден — дальше не “достраиваем”
    If IsAbsolutePath(t) Then ResolveFilePath = "": Exit Function

    Dim baseDir As String: baseDir = ThisWorkbook.path
    Dim cand As String

    ' 2) рядом с книгой
    If Len(baseDir) > 0 Then
        cand = baseDir & "\" & t
        If FileExists(cand) Then ResolveFilePath = cand: Exit Function

        ' 3) типовые подпапки
        cand = baseDir & "\files\" & t
        If FileExists(cand) Then ResolveFilePath = cand: Exit Function

        cand = baseDir & "\docs\" & t
        If FileExists(cand) Then ResolveFilePath = cand: Exit Function

        cand = baseDir & "\attachments\" & t
        If FileExists(cand) Then ResolveFilePath = cand: Exit Function

        cand = baseDir & "\Документы\" & t
        If FileExists(cand) Then ResolveFilePath = cand: Exit Function
    End If

    ' 4) Desktop
    cand = DesktopDir() & "\" & t
    If FileExists(cand) Then ResolveFilePath = cand: Exit Function

    ResolveFilePath = ""
End Function


Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(p, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
    On Error GoTo 0
End Function


Private Function FileNameFromPath(ByVal p As String) As String
    Dim i As Long: i = InStrRev(p, "\")
    If i > 0 Then FileNameFromPath = Mid$(p, i + 1) Else FileNameFromPath = p
End Function

Private Function ReadFileBase64(ByVal p As String) As String
    On Error GoTo Fail
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open: stm.LoadFromFile p
    Dim bytes() As Byte: bytes = stm.Read
    stm.Close
        Dim xml As Object: Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Dim node As Object: Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    ReadFileBase64 = Replace$(node.Text, vbLf, "")
    Exit Function
Fail:
    ReadFileBase64 = ""
End Function

' "Мусор" в ячейках с файлами: Нет/Да/— и т.п.
Private Function IsNoiseFileToken(ByVal token As String) As Boolean
    Dim t As String: t = Norm(token)
    If t = "" Then IsNoiseFileToken = True: Exit Function

    Select Case t
        Case "нет", "да", "не требуется", "отсутствует", "нет данных", "-", "—", "n/a"
            IsNoiseFileToken = True
        Case Else
            IsNoiseFileToken = False
    End Select
End Function


Private Sub ReadFilesIntoObject( _
    ByVal rawCell As Variant, _
    ByVal target As Object, _
    ByVal singleKey As String, _
    ByVal pluralKey As String, _
    Optional ByVal asNameBase64Object As Boolean = False, _
    Optional ByVal nameKey As String = "name" _
)
    Dim s As String: s = Trim$(CStr(rawCell))
    If Len(s) = 0 Then Exit Sub

    Dim tokens As Collection: Set tokens = SplitMany(s)
    If tokens Is Nothing Then Exit Sub

    ' чистим "Нет/Да/—" и прочий мусор
    Dim clean As New Collection
    Dim i As Long, tok As String
    For i = 1 To tokens.Count
        tok = Trim$(CStr(tokens(i)))
        If Len(tok) > 0 Then
            If Not IsNoiseFileToken(tok) Then clean.Add tok
        End If
    Next i
    If clean.Count = 0 Then Exit Sub

    Dim resolved As String, pUse As String, fileLabel As String

    If clean.Count = 1 Then
        tok = CStr(clean(1))

        resolved = ResolveFilePath(tok)
        If Len(resolved) > 0 Then
            pUse = resolved
        Else
            pUse = tok ' хотя бы имя в JSON
        End If

        fileLabel = FileNameFromPath(pUse)

        If asNameBase64Object Then
            Dim o1 As Object: Set o1 = NewDict()
            SafePut o1, nameKey, fileLabel

            If Len(resolved) > 0 Then
                Dim fsz As Long
                On Error Resume Next: fsz = FileLen(resolved): On Error GoTo 0
                If Not (MAX_BASE64_SIZE > 0 And fsz > MAX_BASE64_SIZE) Then
                    SafePut o1, "base64", ReadFileBase64(resolved)
                End If
            End If

            If o1.Count > 0 Then SafePut target, singleKey, o1
        Else
            ' ключ = только имя файла (без локального пути)
            SafePut target, singleKey, fileLabel

            If Len(resolved) > 0 Then
                SafePut target, singleKey & "FileName", fileLabel

                Dim fsz2 As Long
                On Error Resume Next: fsz2 = FileLen(resolved): On Error GoTo 0
                If Not (MAX_BASE64_SIZE > 0 And fsz2 > MAX_BASE64_SIZE) Then
                    SafePut target, singleKey & "Base64", ReadFileBase64(resolved)
                End If
            End If
        End If

    Else
        ' много файлов -> массив объектов
        Dim arr As New Collection
        For i = 1 To clean.Count
            tok = CStr(clean(i))

            resolved = ResolveFilePath(tok)
            If Len(resolved) > 0 Then
                pUse = resolved
            Else
                pUse = tok
            End If

            fileLabel = FileNameFromPath(pUse)

            Dim o As Object: Set o = NewDict()
            If asNameBase64Object Then
                SafePut o, nameKey, fileLabel
            Else
                SafePut o, "fileName", fileLabel
            End If

            If Len(resolved) > 0 Then
                Dim sz As Long
                On Error Resume Next: sz = FileLen(resolved): On Error GoTo 0
                If Not (MAX_BASE64_SIZE > 0 And sz > MAX_BASE64_SIZE) Then
                    SafePut o, "base64", ReadFileBase64(resolved)
                End If
            End If

            If o.Count > 0 Then arr.Add o
        Next i

        If arr.Count > 0 Then SafePut target, pluralKey, arr
    End If
End Sub



' ---------- JSON ----------
Private Function ConvertCollectionToJson(ByVal col As Collection) As String
    Dim json As String, v As Variant
    For Each v In col
        json = json & ConvertItemToJSON(v) & ","
    Next v
    If Len(json) > 0 Then json = Left$(json, Len(json) - 1)
    ConvertCollectionToJson = json
End Function

Private Function ConvertItemToJSON(ByVal item As Variant) As String
    Dim json As String
    If IsObject(item) Then
        Select Case TypeName(item)
            Case "Dictionary"
                json = ConvertDictToJson(item)
            Case "Collection"
                json = ConvertCollectionToJson_sub(item)
            Case Else
                json = """" & EscapeJSON(CStr(item)) & """"
        End Select
    Else
        Select Case VarType(item)
            Case vbEmpty, vbNull
                json = "null"
            Case vbString
                json = """" & EscapeJSON(CStr(item)) & """"
            Case vbDate
                json = """" & Format$(item, "yyyy-mm-dd") & """"
            Case vbBoolean
                json = LCase$(CStr(item))
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                Dim ds As String, sNum As String
                ds = Application.International(xlDecimalSeparator)
                sNum = CStr(item)
                If ds <> "." Then sNum = Replace$(sNum, ds, ".")
                json = sNum
            Case Else
                json = """" & EscapeJSON(CStr(item)) & """"
        End Select
    End If
    ConvertItemToJSON = json
End Function

Private Function ConvertDictToJson(ByVal dict As Object) As String
    Dim json As String, k As Variant
    json = "{"
    For Each k In dict.keys
        json = json & """" & EscapeJSON(CStr(k)) & """:" & ConvertItemToJSON(dict(k)) & ","
    Next k
    If Right$(json, 1) = "," Then json = Left$(json, Len(json) - 1)
    json = json & "}"
    ConvertDictToJson = json
End Function

Private Function ConvertCollectionToJson_sub(ByVal col As Collection) As String
    Dim json As String, v As Variant
    json = "["
    For Each v In col
        json = json & ConvertItemToJSON(v) & ","
    Next v
    If Len(json) > 1 Then json = Left$(json, Len(json) - 1)
    json = json & "]"
    ConvertCollectionToJson_sub = json
End Function

Private Function EscapeJSON(ByVal value As String) As String
    Dim s As String
    s = CStr(value)
    s = Replace$(s, "\", "\\")
    s = Replace$(s, """", "\""")
    s = Replace$(s, "/", "\/")
    s = Replace$(s, vbCrLf, "\n")
    s = Replace$(s, vbCr, "\n")
    s = Replace$(s, vbLf, "\n")
    s = Replace$(s, vbTab, "\t")
    EscapeJSON = s
End Function


' ---------- СОХРАНЕНИЕ UTF-8 ----------
Private Sub SaveUtf8ToDesktop(ByVal fileName As String, ByVal content As String)
    Dim path As String: path = Environ$("USERPROFILE") & "\Desktop\" & fileName
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content
    stm.SaveToFile path, 2
    stm.Close
End Sub

' ===== ДОБАВЛЕНО: утилиты для чанков =====
Private Function FileBaseOnly(ByVal fn As String) As String
    Dim p As Long: p = InStrRev(fn, ".")
    If p > 0 Then FileBaseOnly = Left$(fn, p - 1) Else FileBaseOnly = fn
End Function

Private Function LeftSized(ByRef arr() As String, ByVal n As Long) As String()
    Dim tmp() As String, i As Long
    ReDim tmp(1 To n)
    For i = 1 To n
        tmp(i) = arr(i)
    Next i
    LeftSized = tmp
End Function

Private Sub SaveJsonChunkDesktop(ByVal baseFileName As String, ByVal partIndex As Long, ByRef buf() As String, ByVal countInBuf As Long)
    Dim jsonText As String
    jsonText = "[" & vbCrLf & Join(LeftSized(buf, countInBuf), "," & vbCrLf) & vbCrLf & "]"
    Dim outName As String
    outName = FileBaseOnly(baseFileName) & "_part" & CStr(partIndex) & ".json"
    SaveUtf8ToDesktop outName, jsonText
End Sub


' =====================================================================
' =============== 2) Реестр заказ-нарядов -> JSON ======================
' =====================================================================
Public Sub ConvertCatchOrdersToJSON()
    Dim ws As Worksheet
    Set ws = FindSheetLike("*Реестр заказ-нарядов*")
    If ws Is Nothing Then
        MsgBox "Не найден лист, похожий на '2. Реестр заказ-нарядов...'", vbCritical
        Exit Sub
    End If

    Dim headerRow As Long
    headerRow = FindHeaderRow(ws, "Название поля")
    If headerRow = 0 Then
        MsgBox "Не найдена строка заголовков ('Название поля' в колонке A).", vbCritical
        Exit Sub
    End If

    ' [IMPROVE] Для 2-го листа тоже строим спецификацию заголовков, как для карточек.
    Dim subblockRow As Long
    subblockRow = headerRow - 1

    Dim spec As Collection
    Set spec = BuildHeaderSpec(ws, headerRow, subblockRow)

    Dim startRow As Long
    startRow = GuessDataStartRow(ws, headerRow)

    Dim lastRow As Long
    lastRow = FindLastRowAny(ws)
    If lastRow < startRow Then
        MsgBox "Нет строк данных ниже " & startRow & ".", vbExclamation
        Exit Sub
    End If

    Dim part As Long: part = 1
    Dim chunkCap As Long: chunkCap = IIf(CHUNK_SIZE > 0, CHUNK_SIZE, 1000000000#)
    Dim rowInChunk As Long: rowInChunk = 0
    Dim buf() As String: ReDim buf(1 To chunkCap)

    Dim r As Long, saved As Long
    For r = startRow To lastRow
        If SKIP_EXAMPLE_ROW Then
            Dim aFirst As String
            aFirst = LCase$(Trim$(CStr(ws.Cells(r, 1).value)))
            If InStr(1, aFirst, "пример", vbTextCompare) > 0 Then GoTo NextR
        End If
        If IsRowEmpty(ws, r, headerRow) Then GoTo NextR

        ' [IMPROVE] Передаём spec, больше не жёстко привязаны к номерам колонок.
        Dim item As Object
        Set item = BuildOrderRow(ws, r, spec)

        If item.Count > 0 Then
            rowInChunk = rowInChunk + 1
            buf(rowInChunk) = ConvertItemToJSON(item)
            If rowInChunk = chunkCap Then
                SaveJsonChunkDesktop OUTPUT_FILENAME, part, buf, rowInChunk
                saved = saved + 1
                part = part + 1
                rowInChunk = 0
            End If
        End If
NextR:
    Next r

    If rowInChunk > 0 Then
        SaveJsonChunkDesktop OUTPUT_FILENAME, part, buf, rowInChunk
        saved = saved + 1
    End If

    If saved = 0 Then
        MsgBox "Данные отсутствуют — файлов не создано.", vbInformation
    Else
        MsgBox "JSON сохранён на Рабочем столе, файлов: " & saved & " (база: " & OUTPUT_FILENAME & ")", vbInformation
    End If
End Sub



' [IMPROVE] BuildOrderRow теперь опирается на spec (Подблок + Название поля),
' а не на жёсткие номера колонок. Добавлены поля по организации-исполнителю отлова.
Private Function BuildOrderRow(ws As Worksheet, _
                               r As Long, _
                               spec As Collection) As Object
    Dim root As Object
    Set root = NewDict()

    ' --- "шапка" заказа (orderInfo) ---
    Dim head As Object
    Set head = NewDict()

    Dim i As Long
    Dim col As Long
    Dim h As String
    Dim s As String

    Dim subOrgUpol As String
    Dim subOrgCatch As String
    Dim subAnimal As String

    subOrgUpol = Norm("Общая информация об уполномоченной организации")
    subOrgCatch = Norm("Общие данные об организации по отлову")
    subAnimal = Norm("Информация о животном без владельца")

    Dim firstAnimalCol As Long
    firstAnimalCol = 0

    For i = 1 To spec.Count
        h = spec(i)("hNorm")
        s = spec(i)("sNorm")
        col = spec(i)("col")

        ' --- Общая информация об уполномоченной организации ---
        If s = subOrgUpol Then
            If h = Norm("Субъект РФ") Then
                AddIfPresent ws, r, col, head, "region"
            ElseIf h = Norm("Муниципальный район/округ, городской округ или внутригородская территория") Then
                ' если в шаблоне этого поля нет, просто не заполнится
                AddIfPresent ws, r, col, head, "municipality"
            ElseIf h = Norm("ОГРН") Then
                ' [IMPROVE] ОГРН уполномоченной организации — только цифры, из .Text
                AddIfPresentDigits ws, r, col, head, "ogrn"
            ElseIf h = Norm("ИНН") Then
                ' [IMPROVE] ИНН уполномоченной организации — только цифры
                AddIfPresentDigits ws, r, col, head, "inn"
            ElseIf Left$(h, Len(Norm("наименование организации, уполномоченной на прием заявлений"))) = _
                   Norm("наименование организации, уполномоченной на прием заявлений") Then
                ' [IMPROVE] отдельно сохраняем наименование уполномоченной организации
                AddIfPresent ws, r, col, head, "authorizedOrgName"
            End If
        End If

        ' --- Общие данные об организации по отлову ---
        If s = subOrgCatch Then
            If Left$(h, Len(Norm("наименование организации по отлову"))) = _
               Norm("наименование организации по отлову") Then
                AddIfPresent ws, r, col, head, "catchOrgName"
            ElseIf h = Norm("огрн") Then
                AddIfPresentDigits ws, r, col, head, "catchOrgOgrn"
            ElseIf h = Norm("инн") Then
                AddIfPresentDigits ws, r, col, head, "catchOrgInn"
            End If
        End If

        ' --- Параметры заявки (не завязаны на Подблок) ---
        If h = Norm("номер заявки (заказ-наряда)") Then
            AddIfPresent ws, r, col, head, "orderNumber"
        ElseIf h = Norm("приоритет заявки") Then
            AddIfPresent ws, r, col, head, "priority"
        ElseIf h = Norm("количество дней на отлов") Then
            AddIfPresent ws, r, col, head, "catchDays"
        End If

        ' --- первая колонка с "Номер животного" (для вычисления смещения блоков по животным) ---
        If firstAnimalCol = 0 Then
            If s = subAnimal And h = Norm("номер животного") Then
                firstAnimalCol = col
            End If
        End If
    Next i

    If head.Count > 0 Then
        SafePut root, "orderInfo", head
    End If

    ' --- Животные в заявке ---
    ' В шаблоне до 10 животных, по 14 колонок на каждое животное.
    If firstAnimalCol > 0 Then
        Dim animals As New Collection
        Dim j As Long
        Dim c0 As Long
        Dim lastCol As Long

        lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column

        For j = 0 To 9
            c0 = firstAnimalCol + j * 14
            If c0 > lastCol Then Exit For

            Dim a As Object
            Set a = NewDict()

            If c0 + 0 <= lastCol Then AddIfPresent ws, r, c0 + 0, a, "number"
            If c0 + 1 <= lastCol Then AddIfPresent ws, r, c0 + 1, a, "kind"
            If c0 + 2 <= lastCol Then AddIfPresent ws, r, c0 + 2, a, "color"
            If c0 + 3 <= lastCol Then AddIfPresent ws, r, c0 + 3, a, "size"
            If c0 + 4 <= lastCol Then AddIfPresent ws, r, c0 + 4, a, "unmotivatedAggression"
            If c0 + 5 <= lastCol Then AddIfPresent ws, r, c0 + 5, a, "aggressionDescription"
            If c0 + 6 <= lastCol Then AddIfPresent ws, r, c0 + 6, a, "clip"
            If c0 + 7 <= lastCol Then AddIfPresent ws, r, c0 + 7, a, "clipColor"
            If c0 + 8 <= lastCol Then AddIfPresent ws, r, c0 + 8, a, "extraInfo"
            If c0 + 9 <= lastCol Then AddIfPresent ws, r, c0 + 9, a, "locationAddress"
            If c0 + 10 <= lastCol Then AddIfPresent ws, r, c0 + 10, a, "locationLandmark"

            ' Фотографии и пояснительные записки (строки с путями/именами файлов)
            If c0 + 11 <= lastCol Then
                ReadFilesIntoObject ws.Cells(r, c0 + 11).value, a, "photo", "photos"
            End If

            If c0 + 12 <= lastCol Then AddIfPresent ws, r, c0 + 12, a, "status"

            If c0 + 13 <= lastCol Then
                ReadFilesIntoObject ws.Cells(r, c0 + 13).value, a, "note", "notes"
            End If

            If a.Count > 0 Then animals.Add a
        Next j

        If animals.Count > 0 Then
            SafePut root, "animals", animals
        End If
    End If

    Set BuildOrderRow = root
End Function


' =====================================================================
' =============== 3) Реестр животных без владельцев ====================
' =====================================================================

Public Sub ConvertStrayAnimalsToJSON()
    Dim ws As Worksheet
    Set ws = FindSheetLike("*Реестр животных без владельц*")
    If ws Is Nothing Then
        MsgBox "Не найден лист с именем, похожим на '3. Реестр животных без владельц...'", vbCritical
        Exit Sub
    End If

    Dim headerRow As Long: headerRow = FindHeaderRow(ws, "Название поля")
    If headerRow = 0 Then
        MsgBox "Не найдена строка заголовков ('Название поля' в колонке A).", vbCritical
        Exit Sub
    End If

    Dim subblockRow As Long
    subblockRow = headerRow - 1

    ' [NEW] Строим спецификацию заголовков (Подблок + Название поля), как для 2-го и 4-го листов
    Dim spec As Collection
    Set spec = BuildHeaderSpec(ws, headerRow, subblockRow)

    Dim startRow As Long: startRow = GuessDataStartRow(ws, headerRow)
    Dim lastRow As Long: lastRow = FindLastRowAny(ws)
    If lastRow < startRow Then
        MsgBox "Нет строк данных ниже " & startRow & ".", vbExclamation
        Exit Sub
    End If

    Dim part As Long: part = 1
    Dim chunkCap As Long: chunkCap = IIf(CHUNK_SIZE > 0, CHUNK_SIZE, 1000000000#)
    Dim rowInChunk As Long: rowInChunk = 0
    Dim buf() As String: ReDim buf(1 To chunkCap)

    Dim r As Long, saved As Long
    For r = startRow To lastRow
        If SKIP_EXAMPLE_ROW Then
            Dim aFirst As String
            aFirst = LCase$(Trim$(CStr(ws.Cells(r, 1).value)))
            If InStr(1, aFirst, "пример", vbTextCompare) > 0 Then GoTo NextR2
        End If
        If IsRowEmpty(ws, r, headerRow) Then GoTo NextR2

        Dim item As Object
        Set item = BuildStrayRow(ws, r, spec)

        If item.Count > 0 Then
            rowInChunk = rowInChunk + 1
            buf(rowInChunk) = ConvertItemToJSON(item)
            If rowInChunk = chunkCap Then
                SaveJsonChunkDesktop OUTPUT_FILENAME_STRAY, part, buf, rowInChunk
                saved = saved + 1
                part = part + 1
                rowInChunk = 0
            End If
        End If
NextR2:
    Next r

    If rowInChunk > 0 Then
        SaveJsonChunkDesktop OUTPUT_FILENAME_STRAY, part, buf, rowInChunk
        saved = saved + 1
    End If

    If saved = 0 Then
        MsgBox "Данные отсутствуют — файлов не создано.", vbInformation
    Else
        MsgBox "JSON сохранён на Рабочем столе, файлов: " & saved & " (база: " & OUTPUT_FILENAME_STRAY & ")", vbInformation
    End If
End Sub



Private Function BuildStrayRow(ws As Worksheet, _
                               r As Long, _
                               spec As Collection) As Object
    Dim obj As Object
    Set obj = NewDict()

    Dim i As Long
    Dim col As Long
    Dim h As String
    Dim s As String

    Dim subOrgUpol As String
    Dim subOrgCatch As String
    Dim subAnimal As String
    Dim subCatchInfo As String

    subOrgUpol = Norm("Общая информация об уполномоченной организации")
    subOrgCatch = Norm("Общие данные об организации по отлову")
    subAnimal = Norm("Информация о животном без владельца")
    subCatchInfo = Norm("Сведения об отлове")

    For i = 1 To spec.Count
        h = spec(i)("hNorm")
        s = spec(i)("sNorm")
        col = spec(i)("col")

        ' ===== FIX: Пояснительная записка может быть размечена в любом подблоке (merge в строке "Подблок") =====
        If InStr(1, h, Norm("пояснительная записка"), vbTextCompare) > 0 Then
            ' единый формат как у photo/catchAct: note + noteFileName + noteBase64
            ReadFilesIntoObject ws.Cells(r, col).value, obj, "note", "notes"
        End If

        ' ===== 1. Общая информация об уполномоченной организации =====
        If s = subOrgUpol Then
            If h = Norm("Субъект РФ") Then
                AddIfPresent ws, r, col, obj, "region"
            ElseIf h = Norm("Муниципальный район/округ, городской округ или внутригородская территория") Then
                ' В текущем шаблоне может и не быть, но на всякий случай поддерживаем
                AddIfPresent ws, r, col, obj, "municipality"
            ElseIf h = Norm("ОГРН") Then
                AddIfPresentDigits ws, r, col, obj, "ogrn"
            ElseIf h = Norm("ИНН") Then
                AddIfPresentDigits ws, r, col, obj, "inn"
            ElseIf Left$(h, Len(Norm("наименование организации, уполномоченной на прием заявлений"))) = _
                   Norm("наименование организации, уполномоченной на прием заявлений") Then
                AddIfPresent ws, r, col, obj, "authorizedOrgName"
            End If
        End If

        ' ===== 2. Общие данные об организации по отлову =====
        If s = subOrgCatch Then
            If Left$(h, Len(Norm("наименование организации по отлову"))) = _
               Norm("наименование организации по отлову") Then
                AddIfPresent ws, r, col, obj, "catchOrgName"
            ElseIf h = Norm("огрн") Then
                AddIfPresentDigits ws, r, col, obj, "catchOrgOgrn"
            ElseIf h = Norm("инн") Then
                AddIfPresentDigits ws, r, col, obj, "catchOrgInn"
            End If
        End If

        ' ===== 3. Информация о животном без владельца =====
        If s = subAnimal Then
            If h = Norm("номер животного") Then
                AddIfPresent ws, r, col, obj, "animalNumber"
            ElseIf h = Norm("вид") Then
                AddIfPresent ws, r, col, obj, "type"
            ElseIf h = Norm("пол") Then
                AddIfPresent ws, r, col, obj, "sex"
            ElseIf h = Norm("окрас") Then
                AddIfPresent ws, r, col, obj, "coloration"
            ElseIf h = Norm("размер") Then
                AddIfPresent ws, r, col, obj, "size"
            ElseIf h = Norm("немотивированная агрессия") Then
                AddIfPresent ws, r, col, obj, "unmotivatedAggression"
            ElseIf h = Norm("описание немотивированной агрессии") Then
                AddIfPresent ws, r, col, obj, "aggressionDescription"
            ElseIf h = Norm("клипса") Then
                AddIfPresent ws, r, col, obj, "clip"
            ElseIf h = Norm("цвет клипсы") Then
                AddIfPresent ws, r, col, obj, "clipColor"
            ElseIf h = Norm("дополнительная информация") Then
                AddIfPresent ws, r, col, obj, "additionalInfo"
            ElseIf h = Norm("адрес местонахождения") Then
                AddIfPresent ws, r, col, obj, "locationAddress"
            ElseIf h = Norm("ориентир для нахождения животного") Then
                AddIfPresent ws, r, col, obj, "locationLandmark"
            ElseIf h = Norm("фотография животного") Then
                ReadFilesIntoObject ws.Cells(r, col).value, obj, "photo", "photos"
            ElseIf h = Norm("статус животного") Then
                AddIfPresent ws, r, col, obj, "animalStatus"
            End If
        End If

        ' ===== 4. Сведения об отлове =====
        If s = subCatchInfo Then
            If h = Norm("номер заявки (заказ-наряда)") Then
                AddIfPresent ws, r, col, obj, "orderNumber"
            ElseIf h = Norm("номер муниципального контракта") Then
                AddIfPresent ws, r, col, obj, "municipalContractNumber"
            ElseIf h = Norm("дата муниципального контракта") Then
                AddIfPresentDate ws, r, col, obj, "municipalContractDate"
            ElseIf h = Norm("статус по заявке (заказ-наряду)") Then
                AddIfPresent ws, r, col, obj, "orderStatus"
            ElseIf h = Norm("дата начала отлова") Then
                AddIfPresentDate ws, r, col, obj, "catchStartDate"
            ElseIf h = Norm("время начала отлова") Then
                AddIfPresentTime ws, r, col, obj, "catchStartTime"
            ElseIf h = Norm("дата завершения отлова") Then
                AddIfPresentDate ws, r, col, obj, "catchEndDate"
            ElseIf h = Norm("время завершения отлова") Then
                AddIfPresentTime ws, r, col, obj, "catchEndTime"
            ElseIf h = Norm("адрес отлова") Then
                AddIfPresent ws, r, col, obj, "catchAddress"
            ElseIf h = Norm("видеозапись отлова") Then
                ReadFilesIntoObject ws.Cells(r, col).value, obj, "catchVideo", "catchVideos"
            ElseIf h = Norm("фамилия имя отчество ловца") Then
                AddIfPresent ws, r, col, obj, "catcherFIO"
            ElseIf h = Norm("номер акта отлова") Then
                AddIfPresent ws, r, col, obj, "catchActNumber"
            ElseIf h = Norm("дата составления акта") Then
                AddIfPresentDate ws, r, col, obj, "catchActDate"
            ElseIf h = Norm("акт отлова") Then
                ReadFilesIntoObject ws.Cells(r, col).value, obj, "catchAct", "catchActFiles"
            End If
        End If
    Next i

    Set BuildStrayRow = obj
End Function


' =====================================================================
' ========== 4) Реестр карточек учёта животных без владельцев ==========
' =====================================================================

Private Function SameNorm(ByVal a As String, ByVal b As String) As Boolean
    SameNorm = (Norm(a) = Norm(b))
End Function

Private Function HasPayload(ByVal o As Object, ByVal anchorKey As String) As Boolean
    If o Is Nothing Then Exit Function
    ' Есть хоть что-то из ключевых полей?
    HasPayload = o.Exists(anchorKey) Or o.Exists("date") Or o.Exists("series") _
                 Or o.Exists("dosage") Or o.Exists("employeeFIO") Or o.Exists("drugName") _
                 Or o.Exists("name")
End Function



Public Sub ConvertAnimalCardsToJSON()
    Dim ws As Worksheet
    Set ws = FindSheetLike("*Реестр карточек учета животн*")
    If ws Is Nothing Then
        MsgBox "Не найден лист с именем, похожим на '4. Реестр карточек учета животн...'", vbCritical
        Exit Sub
    End If

    Dim headerRow As Long: headerRow = FindHeaderRow(ws, "Название поля")
    If headerRow = 0 Then
        MsgBox "Не найдена строка 'Название поля' в колонке A.", vbCritical
        Exit Sub
    End If

    Dim subblockRow As Long: subblockRow = headerRow - 1
    Dim startRow As Long: startRow = GuessDataStartRow(ws, headerRow)
    Dim lastRow As Long: lastRow = FindLastRowAny(ws)
    If lastRow < startRow Then
        MsgBox "Нет строк данных ниже " & startRow & ".", vbExclamation
        Exit Sub
    End If

    Dim spec As Collection
    Set spec = BuildHeaderSpec(ws, headerRow, subblockRow)

    Dim part As Long: part = 1
    Dim chunkCap As Long: chunkCap = IIf(CHUNK_SIZE > 0, CHUNK_SIZE, 1000000000#)
    Dim rowInChunk As Long: rowInChunk = 0
    Dim buf() As String: ReDim buf(1 To chunkCap)

    Dim r As Long, saved As Long
    For r = startRow To lastRow
        If SKIP_EXAMPLE_ROW Then
            Dim aFirst As String
            aFirst = LCase$(Trim$(CStr(ws.Cells(r, 1).value)))
            If InStr(1, aFirst, "пример", vbTextCompare) > 0 Then GoTo NextR
        End If
        If IsRowEmpty(ws, r, headerRow) Then GoTo NextR

        Dim item As Object: Set item = BuildCardRow(ws, r, spec)
        If item.Count > 0 Then
            rowInChunk = rowInChunk + 1
            buf(rowInChunk) = ConvertItemToJSON(item)
            If rowInChunk = chunkCap Then
                SaveJsonChunkDesktop OUTPUT_FILENAME_CARDS, part, buf, rowInChunk
                saved = saved + 1
                part = part + 1
                rowInChunk = 0
            End If
        End If
NextR:
    Next r

    If rowInChunk > 0 Then
        SaveJsonChunkDesktop OUTPUT_FILENAME_CARDS, part, buf, rowInChunk
        saved = saved + 1
    End If

    If saved = 0 Then
        MsgBox "Данные отсутствуют — файлов не создано.", vbInformation
    Else
        MsgBox "JSON сохранён на Рабочем столе, файлов: " & saved & " (база: " & OUTPUT_FILENAME_CARDS & ")", vbInformation
    End If
End Sub


Private Function MergedText(ByVal ws As Worksheet, ByVal row As Long, ByVal col As Long) As String
    Dim c As Range
    Set c = ws.Cells(row, col)
    If c.MergeCells Then
        MergedText = CStr(c.MergeArea.Cells(1, 1).value)
    Else
        MergedText = CStr(c.value)
    End If
End Function


' ---------- спецификация заголовков ----------
Private Function BuildHeaderSpec(ws As Worksheet, headerRow As Long, subblockRow As Long) As Collection
    Dim res As New Collection
    Dim lastCol As Long, c As Long
    Dim sRaw As String, hRaw As String, sFill As String
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    sFill = ""
    For c = 2 To lastCol
        ' читаем с учётом merge
        sRaw = MergedText(ws, subblockRow, c)
        hRaw = MergedText(ws, headerRow, c)
        ' если пусто — тащим последний непустой подблок влево-вправо
        If Len(Trim$(sRaw)) = 0 Then
            sRaw = sFill
        Else
            sFill = sRaw
        End If
        
        Dim it As Object: Set it = NewDict()
        SafePut it, "col", c
        SafePut it, "hRaw", hRaw
        SafePut it, "sRaw", sRaw
        SafePut it, "hNorm", Norm(hRaw)
        SafePut it, "sNorm", Norm(sRaw)
        res.Add it
    Next c
    
    Set BuildHeaderSpec = res
End Function


' ---------- сборка одной карточки ----------
Private Function BuildCardRow(ws As Worksheet, r As Long, spec As Collection) As Object
    Dim root As Object: Set root = NewDict()

    ' Базовая карта заголовок -> ключ (без ОГРН/ИНН — их обрабатываем по подблокам)
    Dim generalMap As Object: Set generalMap = NewDict()
    SafePut generalMap, Norm("Субъект РФ"), "region"
    SafePut generalMap, Norm("Муниципальный район/округ, городской округ или внутригородская территория"), "municipality"
    SafePut generalMap, Norm("Номер карточки учета животного без владельца"), "cardNumber"
    SafePut generalMap, Norm("Вид"), "type"
    SafePut generalMap, Norm("Пол"), "sex"
    SafePut generalMap, Norm("Порода"), "breed"
    SafePut generalMap, Norm("Окрас"), "coloration"
    SafePut generalMap, Norm("Размер"), "size"
    SafePut generalMap, Norm("Кличка"), "nickname"
    SafePut generalMap, Norm("Шерсть"), "fur"
    SafePut generalMap, Norm("Уши"), "ears"
    SafePut generalMap, Norm("Хвост"), "tail"
    SafePut generalMap, Norm("Особые приметы"), "specialMarks"
    SafePut generalMap, Norm("Возраст"), "age"
    SafePut generalMap, Norm("Вес"), "weight"
    SafePut generalMap, Norm("Температура"), "temperature"
    SafePut generalMap, Norm("Сведения о нанесенных животному травмах"), "injuriesInfo"
    SafePut generalMap, Norm("Номер метки"), "numMarker"
    SafePut generalMap, Norm("Способ нанесения"), "methodMarker"
    SafePut generalMap, Norm("Статус животного"), "animalStatus"
    SafePut generalMap, Norm("№ клетки, вольера"), "cageNumber"
    SafePut generalMap, Norm("Дата выпуска животного из приюта"), "releaseFromShelterDate"
    SafePut generalMap, Norm("Дата выпуска животного из пункта временного содержания"), "releaseFromPVSDate"
    SafePut generalMap, Norm("Дата, до которой животное содержится на карантине"), "quarantineUntilDate"

    Dim subOrgUpol As String
    Dim subShelterInfo As String
    subOrgUpol = Norm("Общая информация об уполномоченной организации")
    subShelterInfo = Norm("Общая информация о приюте / ПВСе")

    Dim i As Long, col As Long, h As String, s As String

    ' ----- общий проход по заголовкам -----
    For i = 1 To spec.Count
        h = spec(i)("hNorm")
        col = spec(i)("col")
        s = spec(i)("sNorm")

        ' 1) Общие поля (без ОГРН/ИНН)
        If generalMap.Exists(h) Then
            Dim key As String: key = generalMap(h)
            If Right$(key, 4) = "Date" Then
                AddIfPresentDate ws, r, col, root, key
            Else
                AddIfPresent ws, r, col, root, key
            End If
        End If

        ' 2) Общая информация об уполномоченной организации
        If s = subOrgUpol Then
            If h = Norm("ОГРН") Then
                AddIfPresentDigits ws, r, col, root, "ogrn"
            ElseIf h = Norm("ИНН") Then
                AddIfPresentDigits ws, r, col, root, "inn"
            ElseIf Left$(h, Len(Norm("наименование организации, уполномоченной на прием заявлений"))) = _
                   Norm("наименование организации, уполномоченной на прием заявлений") Then
                AddIfPresent ws, r, col, root, "authorizedOrgName"
            End If
        End If

        ' 3) Общая информация о приюте / ПВСе
        If s = subShelterInfo Then
            If Left$(h, Len(Norm("наименование приюта / пункта временного содержания"))) = _
               Norm("наименование приюта / пункта временного содержания") Then
                AddIfPresent ws, r, col, root, "shelterName"
            ElseIf h = Norm("ОГРН") Then
                AddIfPresentDigits ws, r, col, root, "shelterOGRN"
            ElseIf h = Norm("ИНН") Then
                AddIfPresentDigits ws, r, col, root, "shelterINN"
            End If
        End If
    Next i

    ' Фото (любой столбец с «фотограф»)
    For i = 1 To spec.Count
        If InStr(spec(i)("hNorm"), "фотограф") > 0 Then
            ReadFilesIntoObject ws.Cells(r, spec(i)("col")).value, root, "photo", "photos"
            Exit For
        End If
    Next i

    ' Идентификационная метка (паспортная часть)
    Dim idMark As Object: Set idMark = NewDict()
    For i = 1 To spec.Count
        If InStr(spec(i)("sNorm"), "идентификационн") > 0 Then
            h = spec(i)("hNorm"): col = spec(i)("col")
            If h = Norm("Номер метки") Then AddIfPresent ws, r, col, idMark, "number"
            If Left$(h, Len(Norm("Способ нанесения"))) = Norm("Способ нанесения") Then AddIfPresent ws, r, col, idMark, "method"
        End If
    Next i
    If idMark.Count > 0 Then SafePut root, "identityMark", idMark

    ' ===== мероприятия (повторяющиеся блоки) =====
    ' --- ДЕГЕЛЬМИНТИЗАЦИЯ ---
    Dim mapD As Object: Set mapD = NewDict()
    SafePut mapD, Norm("Название препарата"), "drugName"
    SafePut mapD, Norm("Дозировка"), "dosage"
    SafePut mapD, Norm("Дата проведения"), "date"
    SafePut mapD, Norm("Фамилия Имя Отчество сотрудника, проводившего мероприятие"), "employeeFIO"
    SafePut mapD, Norm("Должность сотрудника, проводившего мероприятие"), "employeePosition"
    SafePut mapD, Norm("Номер акта дегельминтизации"), "actNumber"

    Dim deworms As Collection
    Set deworms = ParseEventGroupStrict(ws, r, spec, "дегельминтизац", mapD, "drugName", "Акт дегельминтизации", "dewormAct")
    If deworms.Count > 0 Then SafePut root, "dewormings", deworms

    ' --- ДЕЗИНСЕКЦИЯ ---
    Dim mapS As Object: Set mapS = NewDict()
    SafePut mapS, Norm("Название препарата"), "drugName"
    SafePut mapS, Norm("Дата проведения"), "date"
    SafePut mapS, Norm("Фамилия Имя Отчество сотрудника, проводившего мероприятие"), "employeeFIO"
    SafePut mapS, Norm("Должность сотрудника, проводившего мероприятие"), "employeePosition"
    SafePut mapS, Norm("Номер акта дезинсекции"), "actNumber"

    Dim disins As Collection
    Set disins = ParseEventGroupStrict(ws, r, spec, "дезинсекц", mapS, "drugName", "Акт дезинсекции", "disinsectionAct")
    If disins.Count > 0 Then SafePut root, "disinsections", disins

    ' --- ВАКЦИНАЦИЯ ---
    Dim mapV As Object: Set mapV = NewDict()
    SafePut mapV, Norm("Название препарата"), "drugName"
    SafePut mapV, Norm("Дата проведения"), "date"
    SafePut mapV, Norm("Серия и номер препарата"), "series"
    SafePut mapV, Norm("Дозировка"), "dosage"
    SafePut mapV, Norm("Фамилия Имя Отчество сотрудника, проводившего мероприятие"), "employeeFIO"
    SafePut mapV, Norm("Должность сотрудника, проводившего мероприятие"), "employeePosition"
    SafePut mapV, Norm("Номер акта вакцинации"), "actNumber"

    Dim vaccs As Collection
    Set vaccs = ParseEventGroupStrict(ws, r, spec, "вакцинац", mapV, "drugName", "Акт вакцинации", "vaccinationAct")
    If vaccs.Count > 0 Then SafePut root, "vaccinations", vaccs

    ' --- СТЕРИЛИЗАЦИЯ (кастрация) ---
    Dim mapSt As Object: Set mapSt = NewDict()
    SafePut mapSt, Norm("Название препарата, используемое при стерилизации"), "drugName"
    SafePut mapSt, Norm("Дата проведения"), "date"
    SafePut mapSt, Norm("Дозировка"), "dose"
    SafePut mapSt, Norm("Фамилия Имя Отчество сотрудника, проводившего мероприятие"), "employeeFIO"
    SafePut mapSt, Norm("Должность сотрудника, проводившего мероприятие"), "employeePosition"
    SafePut mapSt, Norm("Номер акта стерилизации (кастрации)"), "actNumber"

    Dim sterArr As Collection
    Set sterArr = ParseEventGroupStrict(ws, r, spec, "стерилизац", mapSt, "drugName", "Акт стерилизации (кастрации)", "sterilizationAct")
    If sterArr.Count > 0 Then SafePut root, "sterilizations", sterArr

    ' — Освидетельствование
    Dim examObj As Object: Set examObj = ParseExamination(ws, r, spec, "освидетельствован")
    If Not examObj Is Nothing Then SafePut root, "examination", examObj

    ' — Эвтаназия
    Dim euthObj As Object: Set euthObj = ParseEuthanasia(ws, r, spec, "эвтаназ")
    If Not euthObj Is Nothing Then SafePut root, "euthanasia", euthObj

    ' — Утилизация
    Dim utilObj As Object: Set utilObj = ParseUtilization(ws, r, spec, "утилизац")
    If Not utilObj Is Nothing Then SafePut root, "utilization", utilObj

    ' — Нанесение идентификационной метки (как мероприятие)
    Dim markDef As Object: Set markDef = NewDict()
    SafePut markDef, Norm("Тип мероприятия"), "type"
    SafePut markDef, Norm("Номер метки"), "number"
    SafePut markDef, Norm("Способ нанесения (клипса, чип, бирка)"), "method"
    SafePut markDef, Norm("Место нанесения"), "place"
    SafePut markDef, Norm("Дата проведения"), "date"
    SafePut markDef, Norm("Фамилия Имя Отчество сотрудника, проводившего мероприятие"), "employeeFIO"
    SafePut markDef, Norm("Должность сотрудника, проводившего мероприятие"), "employeePosition"

    Dim markArr As Collection
    Set markArr = ParseEventGroupStrict(ws, r, spec, "нанесение идентификационн", markDef, "number", "", "")
    If markArr.Count > 0 Then SafePut root, "markingEvents", markArr

    ' — Иные мероприятия
    Dim mapO As Object: Set mapO = NewDict()
    SafePut mapO, Norm("Тип мероприятия"), "type"
    SafePut mapO, Norm("Наименование мероприятия"), "name"
    SafePut mapO, Norm("Описание мероприятия"), "description"
    SafePut mapO, Norm("Дата проведения"), "date"
    SafePut mapO, Norm("Фамилия Имя Отчество сотрудника, проводившего мероприятие"), "employeeFIO"
    SafePut mapO, Norm("Должность сотрудника, проводившего мероприятие"), "employeePosition"
    SafePut mapO, Norm("Номер документа иного мероприятия"), "documentNumber"

    Dim others As Collection
    Set others = ParseEventGroupStrict(ws, r, spec, "иное мероприяти", mapO, "name", "Документ", "otherEventDocument")
    If others.Count > 0 Then SafePut root, "otherEvents", others

    ' ===== справочно-актная часть =====
    Dim relObj As Object: Set relObj = ParseRelease(ws, r, spec, "информация о выпуске животного")
    If Not relObj Is Nothing Then SafePut root, "releaseInfo", relObj

    Dim trObj As Object: Set trObj = ParseTransferOwner(ws, r, spec, "информация о передаче животного")
    If Not trObj Is Nothing Then SafePut root, "transferToOwner", trObj

    Dim deathObj As Object: Set deathObj = ParseDeath(ws, r, spec, "информация о падеже")
    If Not deathObj Is Nothing Then SafePut root, "deathInfo", deathObj

    Dim catchActObj As Object:
    Set catchActObj = ParseHandoverCatcher(ws, r, spec, _
    "Информация об акте приема-передачи животного без владельца с ловцом")
    If Not catchActObj Is Nothing Then SafePut root, "handoverWithCatcher", catchActObj

    Dim shelterActObj As Object
    Set shelterActObj = ParseHandoverShelter(ws, r, spec, _
    "Информация об акте приема-передачи животного без владельца между ПВСом и приютом")


    If Not shelterActObj Is Nothing Then SafePut root, "handoverWithShelter", shelterActObj

    Set BuildCardRow = root
End Function


' ---------- общий парсер групп мероприятий ----------
Private Function ParseEventGroup(ws As Worksheet, r As Long, spec As Collection, groupLike As String, fieldMap As Object, fileHeaderNorm As String, fileJsonBase As String) As Collection
    Dim res As New Collection
    Dim i As Long, col As Long, h As String, s As String
    Dim cur As Object: Set cur = Nothing
    Dim currentType As String  ' запоминаем "Тип мероприятия", чтобы проставлять его в следующих итемах

    For i = 1 To spec.Count
        s = spec(i)("sNorm")
        If Not IsGroupMatch(s, groupLike) Then GoTo NextI

        h = spec(i)("hNorm"): col = spec(i)("col")

        ' Явный старт нового элемента по "Тип мероприятия"
        If h = Norm("Тип мероприятия") Then
            If Not cur Is Nothing Then
                If cur.Count > 0 Then res.Add cur
            End If
            Set cur = NewDict()
            AddIfPresent ws, r, col, cur, "type"
            On Error Resume Next
            currentType = CStr(cur("type"))
            On Error GoTo 0
            GoTo MaybeFiles
        End If

        ' Если ещё не было cur (например, "Тип мероприятия" отсутствует в текущем повторе),
        ' открываем новый и проставляем сохранённый type.
        If cur Is Nothing Then
            Set cur = NewDict()
            If Len(currentType) > 0 Then SafePut cur, "type", currentType
        End If

        ' Основное сопоставление полей по карте
        If fieldMap.Exists(h) Then
            Dim key As String: key = CStr(fieldMap(h))

            ' Если встретили поле, которое логически начало НОВОГО повторяющегося блока
            ' (обычно "Название препарата" / "Наименование мероприятия" / "Серия..." / "Дата проведения"),
            ' а в текущем итеме оно уже было — закрываем предыдущий и начинаем новый.
            If cur.Exists(key) Then
                If key = "drugName" Or key = "name" Or key = "series" Or key = "date" Then
                    If cur.Count > 0 Then res.Add cur
                    Set cur = NewDict()
                    If Len(currentType) > 0 Then SafePut cur, "type", currentType
                End If
            End If

            ' Сама запись поля
            If Right$(LCase$(key), 4) = "date" Then
                AddIfPresentDate ws, r, col, cur, key
            Else
                AddIfPresent ws, r, col, cur, key
            End If
        End If

MaybeFiles:
        ' Файлы акта в текущий итем
        If Len(fileHeaderNorm) > 0 And h = fileHeaderNorm Then
            ReadFilesIntoObject ws.Cells(r, col).value, cur, fileJsonBase, fileJsonBase & "s"
        End If

NextI:
    Next i

    ' Добиваем хвост
    If Not cur Is Nothing Then
        If cur.Count > 0 Then res.Add cur
    End If

    Set ParseEventGroup = res
End Function


' ---------- частные блоки ----------
Private Function ParseExamination(ws As Worksheet, r As Long, spec As Collection, groupLike As String) As Object
    Dim obj As Object: Set obj = NewDict()
    Dim i As Long, col As Long, h As String
    For i = 1 To spec.Count
        If InStr(spec(i)("sNorm"), groupLike) = 0 Then GoTo NextI
        h = spec(i)("hNorm"): col = spec(i)("col")

        If h = Norm("Тип мероприятия") Then AddIfPresent ws, r, col, obj, "type"
        If h = Norm("Реакция на еду в присутствии чужого человека") Then AddIfPresent ws, r, col, obj, "foodReactionPresence"
        If h = Norm("Реакция на еду, предложенную чужим человеком") Then AddIfPresent ws, r, col, obj, "foodReactionOffer"
        If h = Norm("Реакция на резкие звуки") Then AddIfPresent ws, r, col, obj, "loudSoundReaction"
        If h = Norm("Решение, принятое комиссией") Then AddIfPresent ws, r, col, obj, "commissionDecision"
        If Left$(h, Len(Norm("Фамилия имя отчество сотрудника комиссии"))) = Norm("Фамилия имя отчество сотрудника комиссии") Then
            AddIfPresent ws, r, col, obj, "commissionMember" & CStr(i)
        End If
        If h = Norm("Работник, составивший акт") Then AddIfPresent ws, r, col, obj, "actAuthor"
        If h = Norm("Дата проведения освидетельствования") Then AddIfPresentDate ws, r, col, obj, "date"
        If h = Norm("Номер акта освидетельствования") Then AddIfPresent ws, r, col, obj, "actNumber"
        If h = Norm("Акт освидетельствования") Then ReadFilesIntoObject ws.Cells(r, col).value, obj, "actFile", "actFiles"
NextI:
    Next i
    If obj.Count = 0 Then Set ParseExamination = Nothing Else Set ParseExamination = obj
End Function

Private Function ParseEuthanasia(ws As Worksheet, r As Long, spec As Collection, groupLike As String) As Object
    Dim obj As Object: Set obj = NewDict()
    Dim i As Long, col As Long, h As String
    For i = 1 To spec.Count
        If InStr(spec(i)("sNorm"), groupLike) = 0 Then GoTo NextI
        h = spec(i)("hNorm"): col = spec(i)("col")

        If h = Norm("Тип мероприятия") Then AddIfPresent ws, r, col, obj, "type"
        If h = Norm("Причина применения процедуры эвтаназии") Then AddIfPresent ws, r, col, obj, "reason"
        If h = Norm("Дата эвтаназии") Then AddIfPresentDate ws, r, col, obj, "date"
        If h = Norm("Время эвтаназии") Then AddIfPresentTime ws, r, col, obj, "time"
        If h = Norm("Способ эвтаназии") Then AddIfPresent ws, r, col, obj, "method"
        If h = Norm("Название препарата") Then AddIfPresent ws, r, col, obj, "drugName"
        If h = Norm("Дозировка") Then AddIfPresent ws, r, col, obj, "dosage"
        If h = Norm("Фамилия Имя Отчество сотрудника, проводившего мероприятие") Then AddIfPresent ws, r, col, obj, "employeeFIO"
        If h = Norm("Должность сотрудника, проводившего мероприятие") Then AddIfPresent ws, r, col, obj, "employeePosition"
        If h = Norm("Номер акта эвтаназии") Then AddIfPresent ws, r, col, obj, "actNumber"
        If h = Norm("Акт эвтаназии") Then ReadFilesIntoObject ws.Cells(r, col).value, obj, "actFile", "actFiles"
NextI:
    Next i
    If obj.Count = 0 Then Set ParseEuthanasia = Nothing Else Set ParseEuthanasia = obj
End Function

Private Function ParseUtilization(ws As Worksheet, r As Long, spec As Collection, groupLike As String) As Object
    Dim obj As Object: Set obj = NewDict()
    Dim i As Long, col As Long, h As String
    For i = 1 To spec.Count
        If InStr(spec(i)("sNorm"), groupLike) = 0 Then GoTo NextI
        h = spec(i)("hNorm"): col = spec(i)("col")

        If h = Norm("Тип мероприятия") Then AddIfPresent ws, r, col, obj, "type"
        If h = Norm("Дата утилизации") Then AddIfPresentDate ws, r, col, obj, "date"
        If h = Norm("Основание для утилизации животного") Then AddIfPresent ws, r, col, obj, "basis"
        If h = Norm("Способ утилизации") Then AddIfPresent ws, r, col, obj, "method"
        If h = Norm("Фамилия Имя Отчество сотрудника, проводившего мероприятие") Then AddIfPresent ws, r, col, obj, "employeeFIO"
        If h = Norm("Должность сотрудника, проводившего мероприятие") Then AddIfPresent ws, r, col, obj, "employeePosition"
        If h = Norm("Номер акта утилизации") Then AddIfPresent ws, r, col, obj, "actNumber"
        If h = Norm("Акт утилизации") Then ReadFilesIntoObject ws.Cells(r, col).value, obj, "actFile", "actFiles"
NextI:
    Next i
    If obj.Count = 0 Then Set ParseUtilization = Nothing Else Set ParseUtilization = obj
End Function

Private Function ParseRelease(ws As Worksheet, r As Long, spec As Collection, groupLike As String) As Object
    Dim o As Object: Set o = NewDict()
    Dim i As Long, col As Long, h As String
    For i = 1 To spec.Count
        If InStr(spec(i)("sNorm"), groupLike) = 0 Then GoTo NextI
        h = spec(i)("hNorm"): col = spec(i)("col")

        If h = Norm("Наименование акта") Then AddIfPresent ws, r, col, o, "actName"
        If h = Norm("Номер акта о выпуске") Then AddIfPresent ws, r, col, o, "actNumber"
        If h = Norm("Дата составления акта") Then AddIfPresentDate ws, r, col, o, "actDate"

        If h = Norm("Наименование приюта") Then AddIfPresent ws, r, col, o, "shelterName"
        If h = Norm("Адрес приюта") Then AddIfPresent ws, r, col, o, "shelterAddress"
        If h = Norm("ИНН приюта") Then AddIfPresentDigits ws, r, col, o, "shelterINN"
        If h = Norm("ОГРН приюта") Then AddIfPresentDigits ws, r, col, o, "shelterOGRN"

        If h = Norm("Наименование пункта временного содержания") Then AddIfPresent ws, r, col, o, "pvsName"
        If h = Norm("Адрес пункта временного содержания") Then AddIfPresent ws, r, col, o, "pvsAddress"
        If h = Norm("ИНН пункта временного содержания") Then AddIfPresentDigits ws, r, col, o, "pvsINN"
        If h = Norm("ОГРН пункта временного содержания") Then AddIfPresentDigits ws, r, col, o, "pvsOGRN"

        If h = Norm("Фамилия, имя, отчество исполнителя по выпуску") Then AddIfPresent ws, r, col, o, "catcherFIO"
        If h = Norm("Адрес выпуска") Then AddIfPresent ws, r, col, o, "releaseAddress"
        If h = Norm("Акт выпуска") Then ReadFilesIntoObject ws.Cells(r, col).value, o, "actFile", "actFiles"
NextI:
    Next i

    If o.Count = 0 Then Set ParseRelease = Nothing Else Set ParseRelease = o
End Function

Private Function ParseTransferOwner(ws As Worksheet, r As Long, spec As Collection, groupLike As String) As Object
    Dim o As Object: Set o = NewDict()
    Dim i As Long, col As Long, h As String
    For i = 1 To spec.Count
        If InStr(spec(i)("sNorm"), groupLike) = 0 Then GoTo NextI
        h = spec(i)("hNorm"): col = spec(i)("col")

        If h = Norm("Наименование акта") Then AddIfPresent ws, r, col, o, "actName"
        If h = Norm("Номер акта о передаче прежнему или новому владельцу") Then AddIfPresent ws, r, col, o, "actNumber"
        If h = Norm("Дата передачи") Then AddIfPresentDate ws, r, col, o, "transferDate"

        If h = Norm("Наименование приюта") Then AddIfPresent ws, r, col, o, "shelterName"
        If h = Norm("Адрес приюта") Then AddIfPresent ws, r, col, o, "shelterAddress"
        If h = Norm("ИНН приюта") Then AddIfPresentDigits ws, r, col, o, "shelterINN"
        If h = Norm("ОГРН приюта") Then AddIfPresentDigits ws, r, col, o, "shelterOGRN"

        If h = Norm("Наименование пункта временного содержания") Then AddIfPresent ws, r, col, o, "pvsName"
        If h = Norm("Адрес пункта временного содержания") Then AddIfPresent ws, r, col, o, "pvsAddress"
        If h = Norm("ИНН пункта временного содержания") Then AddIfPresentDigits ws, r, col, o, "pvsINN"
        If h = Norm("ОГРН пункта временного содержания") Then AddIfPresentDigits ws, r, col, o, "pvsOGRN"


        If h = Norm("Фамилия, имя, отчество нового владельца") Then AddIfPresent ws, r, col, o, "newOwnerFIO"
        If h = Norm("Адрес") Then AddIfPresent ws, r, col, o, "newOwnerAddress"
        If h = Norm("Серия документа, удостоверяющего личность") Then AddIfPresent ws, r, col, o, "idSeries"
        If h = Norm("Номер документа, удостоверяющего личность") Then AddIfPresent ws, r, col, o, "idNumber"
        If h = Norm("Код подразделения") Then AddIfPresent ws, r, col, o, "idDeptCode"
        If h = Norm("Дата выдачи документа") Then AddIfPresentDate ws, r, col, o, "idIssueDate"
        If h = Norm("Кем выдан") Then AddIfPresent ws, r, col, o, "idIssuedBy"

        If h = Norm("Акт передачи животного прежнему или новому владельцу") Then ReadFilesIntoObject ws.Cells(r, col).value, o, "actFile", "actFiles"
NextI:
    Next i
    If o.Count = 0 Then Set ParseTransferOwner = Nothing Else Set ParseTransferOwner = o
End Function

Private Function ParseDeath(ws As Worksheet, r As Long, spec As Collection, groupLike As String) As Object
    Dim o As Object: Set o = NewDict()
    Dim i As Long, col As Long, h As String
    For i = 1 To spec.Count
        If InStr(spec(i)("sNorm"), groupLike) = 0 Then GoTo NextI
        h = spec(i)("hNorm"): col = spec(i)("col")

        If h = Norm("Наименование акта") Then AddIfPresent ws, r, col, o, "actName"
        If h = Norm("Номер акта о падеже") Then AddIfPresent ws, r, col, o, "actNumber"
        If h = Norm("Дата составления акта") Then AddIfPresentDate ws, r, col, o, "actDate"
        If h = Norm("Дата падежа") Then AddIfPresentDate ws, r, col, o, "deathDate"

        If h = Norm("Наименование приюта") Then AddIfPresent ws, r, col, o, "shelterName"
        If h = Norm("Адрес приюта") Then AddIfPresent ws, r, col, o, "shelterAddress"
        If h = Norm("ИНН приюта") Then AddIfPresentDigits ws, r, col, o, "shelterINN"
        If h = Norm("ОГРН приюта") Then AddIfPresentDigits ws, r, col, o, "shelterOGRN"

        If h = Norm("Наименование пункта временного содержания") Then AddIfPresent ws, r, col, o, "pvsName"
        If h = Norm("Адрес пункта временного содержания") Then AddIfPresent ws, r, col, o, "pvsAddress"
        If h = Norm("ИНН пункта временного содержания") Then AddIfPresentDigits ws, r, col, o, "pvsINN"
        If h = Norm("ОГРН пункта временного содержания") Then AddIfPresentDigits ws, r, col, o, "pvsOGRN"


        If h = Norm("Акт падежа") Then ReadFilesIntoObject ws.Cells(r, col).value, o, "actFile", "actFiles"
NextI:
    Next i
    If o.Count = 0 Then Set ParseDeath = Nothing Else Set ParseDeath = o
End Function

Private Function ParseHandoverCatcher(ws As Worksheet, r As Long, spec As Collection, groupLike As String) As Object
    Dim o As Object: Set o = NewDict()
    Dim i As Long, col As Long, h As String

    For i = 1 To spec.Count
        ' было: If InStr(spec(i)("sNorm"), groupLike) = 0 Then GoTo NextI
        If Not SubblockMatch(CStr(spec(i)("sNorm")), groupLike) Then GoTo NextI

        h = CStr(spec(i)("hNorm")): col = CLng(spec(i)("col"))

        ' NEW: Наименование акта (в твоем примере оно есть)
        If h = Norm("Наименование акта") Then AddIfPresent ws, r, col, o, "actName"

        If h = Norm("Номер акта приема-передачи") Then AddIfPresent ws, r, col, o, "actNumber"
        If h = Norm("Номер заявки (заказ-наряда)") Then AddIfPresent ws, r, col, o, "orderNumber"
        If h = Norm("Дата создания заявки (заказ-наряда)") Then AddIfPresentDate ws, r, col, o, "orderCreateDate"
        If h = Norm("Фамилия, имя, отчество ловца") Then AddIfPresent ws, r, col, o, "catcherFIO"
        If h = Norm("Номер телефона ловца") Then AddIfPresent ws, r, col, o, "catcherPhone"

        If h = Norm("Наименование приюта") Then AddIfPresent ws, r, col, o, "shelterName"
        If h = Norm("Адрес приюта") Then AddIfPresent ws, r, col, o, "shelterAddress"
        If h = Norm("Номер телефона приюта") Then AddIfPresent ws, r, col, o, "shelterPhone"

        If h = Norm("Наименование пункта временного содержания") Then AddIfPresent ws, r, col, o, "pvsName"
        If h = Norm("Адрес пункта временного содержания") Then AddIfPresent ws, r, col, o, "pvsAddress"
        If h = Norm("Номер телефона пункта временного содержания") Then AddIfPresent ws, r, col, o, "pvsPhone"

        ' Файл: делаем чуть устойчивее, чтобы не сломалось из-за вариаций заголовка
        If h = Norm("Акт приема-передачи") _
           Or (InStr(h, Norm("акт приема-передачи")) > 0 _
               And InStr(h, Norm("номер")) = 0 _
               And InStr(h, Norm("дата")) = 0) Then
            ReadFilesIntoObject ws.Cells(r, col).value, o, "actFile", "actFiles"
        End If
NextI:
    Next i

    If o.Count = 0 Then Set ParseHandoverCatcher = Nothing Else Set ParseHandoverCatcher = o
End Function


Private Function ParseHandoverShelter(ws As Worksheet, r As Long, spec As Collection, groupLike As String) As Object
    Dim o As Object: Set o = NewDict()
    Dim i As Long, col As Long, h As String

    For i = 1 To spec.Count
        If Not SubblockMatch(CStr(spec(i)("sNorm")), groupLike) Then GoTo NextI

        h = CStr(spec(i)("hNorm")): col = CLng(spec(i)("col"))

        If h = Norm("Наименование акта") Then AddIfPresent ws, r, col, o, "actName"
        If h = Norm("Номер акта приема-передачи") Then AddIfPresent ws, r, col, o, "actNumber"
        If h = Norm("Дата составления акта") Then AddIfPresentDate ws, r, col, o, "actDate"
        If h = Norm("Наименование пункта временного содержания") Then AddIfPresent ws, r, col, o, "pvsName"
        If h = Norm("Наименование приюта") Then AddIfPresent ws, r, col, o, "shelterName"

        If h = Norm("Акт приема-передачи") _
           Or (InStr(h, Norm("акт приема-передачи")) > 0 _
               And InStr(h, Norm("номер")) = 0 _
               And InStr(h, Norm("дата")) = 0) Then
            ReadFilesIntoObject ws.Cells(r, col).value, o, "actFile", "actFiles"
        End If
NextI:
    Next i

    If o.Count = 0 Then Set ParseHandoverShelter = Nothing Else Set ParseHandoverShelter = o
End Function

