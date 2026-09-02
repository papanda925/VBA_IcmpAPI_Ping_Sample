Attribute VB_Name = "TraceLogger"
Option Explicit

' ============================================================================
' TraceLogger.bas
' ----------------------------------------------------------------------------
' 通信処理の「いつ・どちら側で・何をしたか」を、次の2か所へ同時に記録する。
'   1. VBEのイミディエイトウィンドウ（Ctrl + Gで表示）
'   2. このマクロを保存したブックの「Trace」シート
'
' 通信本体からログ処理を分離することで、ICMP APIの学習に集中しやすくする。
' Traceシートを作成できない場合でも、通信処理自体は止めない設計としている。
' ============================================================================

Private mTraceSequence As Long

' Traceシートを準備する。
' clearExisting=True の場合は、以前の記録を消して見出しから作り直す。
Public Sub InitializeTraceSheet(Optional ByVal clearExisting As Boolean = True)

    Dim ws As Worksheet
    Dim headers As Variant
    Dim columnIndex As Long

    On Error GoTo SheetUnavailable

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Trace")
    On Error GoTo SheetUnavailable

    If ws Is Nothing Then
        ' Traceシートがない場合だけ、新しいワークシートを末尾へ追加する。
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Trace"
    ElseIf clearExisting Then
        ws.Cells.Clear
    End If

    headers = Array( _
        "Time", "TraceId", "Side", "Step", _
        "Direction", "Detail", "Result", "ElapsedMs")

    For columnIndex = LBound(headers) To UBound(headers)
        ws.Cells(1, columnIndex + 1).Value2 = headers(columnIndex)
    Next

    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With

    ' 日時、TraceId、方向、説明、結果をExcelが日付・指数・数式へ自動変換
    ' しないよう、文字列として扱う。H列の経過時間だけは数値のまま残す。
    ws.Columns("A:G").NumberFormat = "@"
    ws.Columns("H").NumberFormat = "0"
    ws.Columns("A:H").AutoFit

    Exit Sub

SheetUnavailable:
    ' シート保護等で準備できなくても、以降はイミディエイト出力だけで継続できる。
    Debug.Print CurrentTimeText() & " | TRACE-WARNING | Traceシートを準備できません: " & _
                Err.Number & " / " & Err.Description

End Sub

' 1回の処理を結び付けるTraceIdを作成する。
' 同じ秒に複数回実行しても区別できるよう、連番を末尾に付ける。
Public Function NewTraceId(ByVal prefix As String) As String

    mTraceSequence = mTraceSequence + 1

    NewTraceId = UCase$(prefix) & "-" & _
                 Format$(Now, "yyyymmdd-hhnnss") & "-" & _
                 Format$(mTraceSequence, "000")

End Function

' Timer関数で取得した開始時刻から、経過ミリ秒を求める。
' Timerは午前0時に0へ戻るため、日付をまたいだ場合だけ1日分を補正する。
Public Function ElapsedMilliseconds(ByVal startedAt As Double) As Long

    Dim elapsedSeconds As Double

    elapsedSeconds = CDbl(Timer) - startedAt
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#

    ElapsedMilliseconds = CLng(elapsedSeconds * 1000#)

End Function

' トレースをイミディエイトウィンドウとTraceシートへ記録する。
' Detailに改行が含まれる場合は1行へ変換し、一覧性を保つ。
Public Sub WriteTrace( _
    ByVal traceId As String, _
    ByVal side As String, _
    ByVal stepName As String, _
    ByVal direction As String, _
    ByVal detail As String, _
    ByVal result As String, _
    ByVal elapsedMs As Long)

    Dim timeText As String
    Dim traceLine As String
    Dim ws As Worksheet
    Dim nextRow As Long

    timeText = CurrentTimeText()
    detail = Replace$(Replace$(detail, vbCr, " "), vbLf, " ")

    traceLine = timeText & " | " & traceId & " | " & side & " | " & _
                stepName & " | " & direction & " | " & detail & " | " & _
                result & " | " & CStr(elapsedMs) & " ms"

    ' Debug.Printは、シートへの出力でエラーが起きても必ず先に実行する。
    Debug.Print traceLine

    On Error GoTo SheetUnavailable

    Set ws = ThisWorkbook.Worksheets("Trace")
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' 既存シートを再利用した場合も、この行のA:Gを文字列形式へ戻してから書く。
    ' Detailが「=」から始まっても数式として評価されないための安全対策である。
    ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 7)).NumberFormat = "@"
    ws.Cells(nextRow, 1).Value2 = timeText
    ws.Cells(nextRow, 2).Value2 = traceId
    ws.Cells(nextRow, 3).Value2 = side
    ws.Cells(nextRow, 4).Value2 = stepName
    ws.Cells(nextRow, 5).Value2 = direction
    ws.Cells(nextRow, 6).Value2 = detail
    ws.Cells(nextRow, 7).Value2 = result
    ws.Cells(nextRow, 8).Value2 = elapsedMs

    Exit Sub

SheetUnavailable:
    ' 読み取り専用ブックや保護されたブックでも、Ping本体は継続させる。
    Debug.Print CurrentTimeText() & " | TRACE-WARNING | Traceシートへ記録できません: " & _
                Err.Number & " / " & Err.Description

End Sub

' 現在時刻を「時:分:秒.ミリ秒」に整形する。
Private Function CurrentTimeText() As String

    Dim timerValue As Double
    Dim milliseconds As Long

    timerValue = CDbl(Timer)
    milliseconds = CLng((timerValue - Fix(timerValue)) * 1000#)
    If milliseconds > 999 Then milliseconds = 999

    CurrentTimeText = Format$(Now, "hh:nn:ss") & "." & _
                      Format$(milliseconds, "000")

End Function
