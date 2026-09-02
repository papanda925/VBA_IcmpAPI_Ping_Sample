Attribute VB_Name = "IcmpPingIPv4"
Option Explicit

' ============================================================================
' IcmpPingIPv4.bas
' ----------------------------------------------------------------------------
' WindowsのICMP APIを直接呼び出し、IPv4 Pingを実行する学習用サンプル。
'
' 【このサンプルで使用しないもの】
'   ・ping.exe
'   ・WScript.Shell / Shell関数
'   ・PowerShell
'
' VBAから iphlpapi.dll の IcmpSendEcho を直接呼ぶため、要求を送ってから応答を
' 読み取るまでの流れを、処理単位で観察できる。
'
' 対象はVBA7を搭載した32bit / 64bit版Excel。LongPtrを使うことで、ハンドルや
' ポインターの大きさがOfficeのビット数に合わせて切り替わる。
' ============================================================================

' ICMP Echo要求を送るためのハンドルを作成する。
Private Declare PtrSafe Function IcmpCreateFile Lib "iphlpapi.dll" () As LongPtr

' ICMPハンドルを閉じる。作成したハンドルは成功・失敗にかかわらず必ず閉じる。
Private Declare PtrSafe Function IcmpCloseHandle Lib "iphlpapi.dll" ( _
    ByVal icmpHandle As LongPtr) As Long

' IPv4宛てにICMP Echo要求を送信し、応答をreplyBufferへ格納する。
' RequestOptions=0を渡すため、このPingサンプルではTTL等の個別指定を行わない。
Private Declare PtrSafe Function IcmpSendEcho Lib "iphlpapi.dll" ( _
    ByVal icmpHandle As LongPtr, _
    ByVal destinationAddress As Long, _
    ByVal requestData As LongPtr, _
    ByVal requestSize As Integer, _
    ByVal requestOptions As LongPtr, _
    ByVal replyBuffer As LongPtr, _
    ByVal replySize As Long, _
    ByVal timeoutMilliseconds As Long) As Long

' バイト配列と数値の間で4バイトをそのままコピーするために使用する。
' 文字列変換ではなくメモリ上の並びを保つことが、IPアドレス処理では重要。
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef destination As Any, _
    ByRef source As Any, _
    ByVal length As LongPtr)

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const IP_SUCCESS As Long = 0
Private Const DEFAULT_TIMEOUT_MS As Long = 1000
Private Const DEFAULT_REQUEST_COUNT As Long = 4
Private Const REPLY_BUFFER_SIZE As Long = 4096

' 初めて試すときの入口。
' 外部ネットワークを使わず、自分自身を表す127.0.0.1へ4回送信する。
Public Sub RunPingLocalhost()

    Dim succeeded As Boolean

    InitializeTraceSheet True
    succeeded = PingIPv4("127.0.0.1", DEFAULT_REQUEST_COUNT, DEFAULT_TIMEOUT_MS)

    If succeeded Then
        MsgBox "IPv4 Pingが完了しました。" & vbCrLf & _
               "Traceシートで処理の流れを確認してください。", _
               vbInformation, "VBA ICMP Ping"
    Else
        MsgBox "IPv4 Pingに成功しませんでした。" & vbCrLf & _
               "TraceシートのERRORまたはTIMEOUTを確認してください。", _
               vbExclamation, "VBA ICMP Ping"
    End If

End Sub

' 指定したドット区切りIPv4アドレスへPingを送る。
' 戻り値は「1回以上IP_SUCCESSの応答を受信したか」を表す。
Public Function PingIPv4( _
    ByVal destinationText As String, _
    Optional ByVal requestCount As Long = DEFAULT_REQUEST_COUNT, _
    Optional ByVal timeoutMilliseconds As Long = DEFAULT_TIMEOUT_MS) As Boolean

    Dim sessionTraceId As String
    Dim requestTraceId As String
    Dim sessionStartedAt As Double
    Dim requestStartedAt As Double
    Dim destinationAddress As Long
    Dim icmpHandle As LongPtr
    Dim payload() As Byte
    Dim replyBuffer() As Byte
    Dim replyCount As Long
    Dim lastDllError As Long
    Dim responseAddress As Long
    Dim responseStatus As Long
    Dim roundTripTime As Long
    Dim responseDataSize As Integer
    Dim requestNumber As Long
    Dim successCount As Long
    Dim closeResult As Long

    On Error GoTo UnexpectedError

    sessionTraceId = NewTraceId("PING4")
    sessionStartedAt = CDbl(Timer)

    WriteTrace sessionTraceId, "CLIENT", "VALIDATE", "LOCAL", _
               "Destination=" & destinationText & _
               ", Count=" & CStr(requestCount) & _
               ", Timeout=" & CStr(timeoutMilliseconds) & "ms", _
               "START", 0

    ' 学習用サンプルから大量送信にならないよう、1回の呼び出しは最大10回に制限する。
    If requestCount < 1 Or requestCount > 10 Then
        WriteTrace sessionTraceId, "CLIENT", "VALIDATE", "LOCAL", _
                   "RequestCount must be between 1 and 10", "INVALID_ARGUMENT", _
                   ElapsedMilliseconds(sessionStartedAt)
        GoTo CleanUp
    End If

    ' 入力ミスで極端に長く待ち続けないよう、上限を60秒にする。
    If timeoutMilliseconds < 1 Or timeoutMilliseconds > 60000 Then
        WriteTrace sessionTraceId, "CLIENT", "VALIDATE", "LOCAL", _
                   "Timeout must be between 1 and 60000ms", "INVALID_ARGUMENT", _
                   ElapsedMilliseconds(sessionStartedAt)
        GoTo CleanUp
    End If

    ' DNS名前解決を混ぜず、教材の焦点をICMP処理へ絞るため、ここでは
    ' 「127.0.0.1」のようなドット区切りIPv4アドレスだけを受け付ける。
    If Not TryParseIPv4(destinationText, destinationAddress) Then
        WriteTrace sessionTraceId, "CLIENT", "PARSE_ADDRESS", "LOCAL", _
                   destinationText, "INVALID_IPV4_ADDRESS", _
                   ElapsedMilliseconds(sessionStartedAt)
        GoTo CleanUp
    End If

    WriteTrace sessionTraceId, "CLIENT", "PARSE_ADDRESS", "LOCAL", _
               destinationText, "OK", ElapsedMilliseconds(sessionStartedAt)

    icmpHandle = IcmpCreateFile()
    If icmpHandle = INVALID_HANDLE_VALUE Then
        lastDllError = Err.LastDllError
        WriteTrace sessionTraceId, "CLIENT", "OPEN_ICMP_HANDLE", "LOCAL", _
                   "IcmpCreateFile", "ERROR=" & CStr(lastDllError), _
                   ElapsedMilliseconds(sessionStartedAt)
        GoTo CleanUp
    End If

    WriteTrace sessionTraceId, "CLIENT", "OPEN_ICMP_HANDLE", "LOCAL", _
               "IcmpCreateFile", "OK", ElapsedMilliseconds(sessionStartedAt)

    ' StrConv(..., vbFromUnicode)でVBAのUnicode文字列を1バイト列へ変換する。
    ' この固定メッセージはASCII文字だけなので、日本語Windowsでも同じ内容になる。
    payload = StrConv("VBA ICMP Ping sample", vbFromUnicode)
    ReDim replyBuffer(0 To REPLY_BUFFER_SIZE - 1)

    For requestNumber = 1 To requestCount

        requestTraceId = sessionTraceId & "-" & Format$(requestNumber, "00")
        requestStartedAt = CDbl(Timer)

        WriteTrace requestTraceId, "CLIENT", "SEND_ECHO", "OUT", _
                   "To=" & destinationText & _
                   ", Bytes=" & CStr(UBound(payload) - LBound(payload) + 1), _
                   "SENDING", 0

        ' VarPtrは、バイト配列の先頭アドレスをAPIへ渡すために使う。
        ' API呼び出し中はpayloadとreplyBufferが有効なままなので、ポインター先も有効。
        replyCount = IcmpSendEcho( _
                        icmpHandle, _
                        destinationAddress, _
                        VarPtr(payload(LBound(payload))), _
                        CInt(UBound(payload) - LBound(payload) + 1), _
                        0, _
                        VarPtr(replyBuffer(0)), _
                        REPLY_BUFFER_SIZE, _
                        timeoutMilliseconds)

        If replyCount = 0 Then
            ' 0は応答を取得できなかったことを示す。API直後にLastDllErrorを保存する。
            lastDllError = Err.LastDllError
            WriteTrace requestTraceId, "TARGET", "RECEIVE_REPLY", "IN", _
                       "No ICMP reply", _
                       IIf(lastDllError = 0, "NO_REPLY", IcmpStatusText(lastDllError)), _
                       ElapsedMilliseconds(requestStartedAt)
        Else
            ' ICMP_ECHO_REPLYの先頭部分はビット数にかかわらず次の並びになる。
            '   0～3 : 応答元IPv4アドレス
            '   4～7 : IPステータス
            '   8～11: 往復時間（ミリ秒）
            '  12～13: 返却データ長
            ' 後半にはポインターがあり32/64bitで大きさが変わるため、教材では
            ' 必要な先頭フィールドだけをバイト位置から安全に読み取る。
            responseAddress = ReadLong(replyBuffer, 0)
            responseStatus = ReadLong(replyBuffer, 4)
            roundTripTime = ReadLong(replyBuffer, 8)
            responseDataSize = ReadInteger(replyBuffer, 12)

            WriteTrace requestTraceId, "TARGET", "RECEIVE_REPLY", "IN", _
                       "From=" & IPv4ToString(responseAddress) & _
                       ", Bytes=" & CStr(responseDataSize) & _
                       ", RTT=" & CStr(roundTripTime) & "ms", _
                       IcmpStatusText(responseStatus), _
                       ElapsedMilliseconds(requestStartedAt)

            If responseStatus = IP_SUCCESS Then successCount = successCount + 1
        End If

        ' 長い連続実行時にもExcelが画面更新やキャンセル操作を処理できるようにする。
        DoEvents

    Next requestNumber

    PingIPv4 = (successCount > 0)

CleanUp:
    If icmpHandle <> 0 And icmpHandle <> INVALID_HANDLE_VALUE Then
        closeResult = IcmpCloseHandle(icmpHandle)
        WriteTrace sessionTraceId, "CLIENT", "CLOSE_ICMP_HANDLE", "LOCAL", _
                   "IcmpCloseHandle", _
                   IIf(closeResult <> 0, "OK", "ERROR=" & CStr(Err.LastDllError)), _
                   ElapsedMilliseconds(sessionStartedAt)
        icmpHandle = 0
    End If

    WriteTrace sessionTraceId, "CLIENT", "SUMMARY", "LOCAL", _
               "Success=" & CStr(successCount) & "/" & CStr(requestCount), _
               IIf(PingIPv4, "COMPLETED", "COMPLETED_WITHOUT_SUCCESS"), _
               ElapsedMilliseconds(sessionStartedAt)
    Exit Function

UnexpectedError:
    WriteTrace sessionTraceId, "CLIENT", "UNEXPECTED_ERROR", "LOCAL", _
               "VBA Error " & CStr(Err.Number) & ": " & Err.Description, _
               "ERROR", ElapsedMilliseconds(sessionStartedAt)
    PingIPv4 = False
    Resume CleanUp

End Function

' ドット区切りIPv4文字列を、IcmpSendEchoが求める4バイト値へ変換する。
Private Function TryParseIPv4( _
    ByVal addressText As String, _
    ByRef addressValue As Long) As Boolean

    Dim parts As Variant
    Dim addressBytes(0 To 3) As Byte
    Dim index As Long
    Dim octetValue As Long

    On Error GoTo InvalidAddress

    parts = Split(Trim$(addressText), ".")
    If UBound(parts) - LBound(parts) + 1 <> 4 Then Exit Function

    For index = 0 To 3
        If Len(parts(index)) = 0 Then Exit Function
        If Not IsNumeric(parts(index)) Then Exit Function

        octetValue = CLng(parts(index))
        If octetValue < 0 Or octetValue > 255 Then Exit Function
        If CStr(octetValue) <> parts(index) And parts(index) <> "0" Then Exit Function

        addressBytes(index) = CByte(octetValue)
    Next index

    ' 例: 127.0.0.1 はメモリ上で 7F 00 00 01 の順に並べる必要がある。
    CopyMemory addressValue, addressBytes(0), 4
    TryParseIPv4 = True
    Exit Function

InvalidAddress:
    TryParseIPv4 = False

End Function

' ICMP_ECHO_REPLYの指定位置から4バイト整数を読み取る。
Private Function ReadLong(ByRef buffer() As Byte, ByVal offset As Long) As Long

    Dim value As Long
    CopyMemory value, buffer(offset), 4
    ReadLong = value

End Function

' ICMP_ECHO_REPLYの指定位置から2バイト整数を読み取る。
Private Function ReadInteger(ByRef buffer() As Byte, ByVal offset As Long) As Integer

    Dim value As Integer
    CopyMemory value, buffer(offset), 2
    ReadInteger = value

End Function

' 4バイトのIPv4値を、人が読みやすいドット区切りへ戻す。
Private Function IPv4ToString(ByVal addressValue As Long) As String

    Dim addressBytes(0 To 3) As Byte

    CopyMemory addressBytes(0), addressValue, 4
    IPv4ToString = CStr(addressBytes(0)) & "." & _
                       CStr(addressBytes(1)) & "." & _
                       CStr(addressBytes(2)) & "." & _
                       CStr(addressBytes(3))

End Function

' WindowsのIP_STATUS値を、トレースで理解しやすい名称へ変換する。
Private Function IcmpStatusText(ByVal statusCode As Long) As String

    Select Case statusCode
        Case 0: IcmpStatusText = "IP_SUCCESS"
        Case 11001: IcmpStatusText = "IP_BUF_TOO_SMALL (11001)"
        Case 11002: IcmpStatusText = "IP_DEST_NET_UNREACHABLE (11002)"
        Case 11003: IcmpStatusText = "IP_DEST_HOST_UNREACHABLE (11003)"
        Case 11004: IcmpStatusText = "IP_DEST_PROT_UNREACHABLE (11004)"
        Case 11005: IcmpStatusText = "IP_DEST_PORT_UNREACHABLE (11005)"
        Case 11006: IcmpStatusText = "IP_NO_RESOURCES (11006)"
        Case 11007: IcmpStatusText = "IP_BAD_OPTION (11007)"
        Case 11008: IcmpStatusText = "IP_HW_ERROR (11008)"
        Case 11009: IcmpStatusText = "IP_PACKET_TOO_BIG (11009)"
        Case 11010: IcmpStatusText = "IP_REQ_TIMED_OUT (11010)"
        Case 11011: IcmpStatusText = "IP_BAD_REQ (11011)"
        Case 11012: IcmpStatusText = "IP_BAD_ROUTE (11012)"
        Case 11013: IcmpStatusText = "IP_TTL_EXPIRED_TRANSIT (11013)"
        Case 11014: IcmpStatusText = "IP_TTL_EXPIRED_REASSEM (11014)"
        Case 11015: IcmpStatusText = "IP_PARAM_PROBLEM (11015)"
        Case 11016: IcmpStatusText = "IP_SOURCE_QUENCH (11016)"
        Case 11017: IcmpStatusText = "IP_OPTION_TOO_BIG (11017)"
        Case 11018: IcmpStatusText = "IP_BAD_DESTINATION (11018)"
        Case Else: IcmpStatusText = "STATUS=" & CStr(statusCode)
    End Select

End Function
