Attribute VB_Name = "IcmpPingIPv6Loopback"
Option Explicit

' ============================================================================
' IcmpPingIPv6Loopback.bas
' ----------------------------------------------------------------------------
' IPv6のループバックアドレス「::1」だけを対象にした最小教材。
' Icmp6CreateFile / Icmp6SendEcho2を直接呼び、ping.exeは使用しない。
'
' 一般のIPv6文字列をSOCKADDR_IN6へ変換する処理は、DNS・スコープID・複数NIC等の
' 別テーマを含む。そのため本モジュールは::1に限定し、IPv6 ICMP APIの基本構造を
' 誤解なく観察できる範囲に絞っている。
' ============================================================================

Private Declare PtrSafe Function Icmp6CreateFile Lib "iphlpapi.dll" () As LongPtr

Private Declare PtrSafe Function IcmpCloseHandle Lib "iphlpapi.dll" ( _
    ByVal icmpHandle As LongPtr) As Long

Private Declare PtrSafe Function Icmp6SendEcho2 Lib "iphlpapi.dll" ( _
    ByVal icmpHandle As LongPtr, _
    ByVal eventHandle As LongPtr, _
    ByVal apcRoutine As LongPtr, _
    ByVal apcContext As LongPtr, _
    ByVal sourceAddress As LongPtr, _
    ByVal destinationAddress As LongPtr, _
    ByVal requestData As LongPtr, _
    ByVal requestSize As Integer, _
    ByVal requestOptions As LongPtr, _
    ByVal replyBuffer As LongPtr, _
    ByVal replySize As Long, _
    ByVal timeoutMilliseconds As Long) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef destination As Any, _
    ByRef source As Any, _
    ByVal length As LongPtr)

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const AF_INET6 As Integer = 23
Private Const SOCKADDR_IN6_SIZE As Long = 28
Private Const REPLY_BUFFER_SIZE As Long = 4096

' IPv6版の入口。送信元・送信先の両方へ::1を設定するため、外部通信は発生しない。
Public Sub RunPingIPv6Loopback()

    Dim traceId As String
    Dim startedAt As Double
    Dim icmpHandle As LongPtr
    Dim sourceAddress(0 To SOCKADDR_IN6_SIZE - 1) As Byte
    Dim destinationAddress(0 To SOCKADDR_IN6_SIZE - 1) As Byte
    Dim payload() As Byte
    Dim replyBuffer(0 To REPLY_BUFFER_SIZE - 1) As Byte
    Dim replyCount As Long
    Dim statusCode As Long
    Dim roundTripTime As Long
    Dim lastDllError As Long
    Dim closeResult As Long
    Dim succeeded As Boolean

    On Error GoTo UnexpectedError

    InitializeTraceSheet True
    traceId = NewTraceId("PING6")
    startedAt = CDbl(Timer)

    WriteTrace traceId, "CLIENT", "BUILD_ADDRESS", "LOCAL", _
               "Source=::1, Destination=::1", "START", 0

    ' SOCKADDR_IN6は28バイト。
    '   0～1 : アドレスファミリ(AF_INET6=23)
    '   2～3 : ポート番号（ICMPでは0）
    '   4～7 : フロー情報（ここでは0）
    '   8～23: 16バイトのIPv6アドレス
    '  24～27: スコープID（::1では0）
    ' ::1は16バイトの最後だけが1なので、配列の23番目を1にする。
    PutInteger sourceAddress, 0, AF_INET6
    sourceAddress(23) = 1
    PutInteger destinationAddress, 0, AF_INET6
    destinationAddress(23) = 1

    WriteTrace traceId, "CLIENT", "BUILD_ADDRESS", "LOCAL", _
               "SOCKADDR_IN6 (28 bytes)", "OK", ElapsedMilliseconds(startedAt)

    icmpHandle = Icmp6CreateFile()
    If icmpHandle = INVALID_HANDLE_VALUE Then
        lastDllError = Err.LastDllError
        WriteTrace traceId, "CLIENT", "OPEN_ICMP6_HANDLE", "LOCAL", _
                   "Icmp6CreateFile", "ERROR=" & CStr(lastDllError), _
                   ElapsedMilliseconds(startedAt)
        GoTo CleanUp
    End If

    WriteTrace traceId, "CLIENT", "OPEN_ICMP6_HANDLE", "LOCAL", _
               "Icmp6CreateFile", "OK", ElapsedMilliseconds(startedAt)

    payload = StrConv("VBA ICMPv6 loopback sample", vbFromUnicode)

    WriteTrace traceId, "CLIENT", "SEND_ECHO", "OUT", _
               "To=::1, Bytes=" & CStr(UBound(payload) + 1), "SENDING", _
               ElapsedMilliseconds(startedAt)

    ' eventHandleとapcRoutineを0にすると同期呼び出しになる。
    ' この教材では非同期完了通知ではなく、IPv6アドレス構造と応答解析に焦点を置く。
    replyCount = Icmp6SendEcho2( _
                    icmpHandle, _
                    0, 0, 0, _
                    VarPtr(sourceAddress(0)), _
                    VarPtr(destinationAddress(0)), _
                    VarPtr(payload(LBound(payload))), _
                    CInt(UBound(payload) - LBound(payload) + 1), _
                    0, _
                    VarPtr(replyBuffer(0)), _
                    REPLY_BUFFER_SIZE, _
                    1000)

    If replyCount = 0 Then
        lastDllError = Err.LastDllError
        WriteTrace traceId, "TARGET", "RECEIVE_REPLY", "IN", _
                   "No ICMPv6 reply", _
                   IIf(lastDllError = 0, "NO_REPLY", "ERROR=" & CStr(lastDllError)), _
                   ElapsedMilliseconds(startedAt)
    Else
        ' ICMPV6_ECHO_REPLYの先頭はpack(1)のIPV6_ADDRESS_EX（26バイト）です。
        ' 親構造体でULONGを4バイト境界へそろえる2バイトのパディングが入り、
        ' 28～31がStatus、32～35がRoundTripTimeとなります。
        ' IPV6_ADDRESS_EX内のアドレスはport(2)+flowinfo(4)の後、offset 6です。
        statusCode = ReadLong(replyBuffer, 28)
        roundTripTime = ReadLong(replyBuffer, 32)
        succeeded = (statusCode = 0)

        WriteTrace traceId, "TARGET", "RECEIVE_REPLY", "IN", _
                   "From=" & ReadIPv6Address(replyBuffer, 6) & _
                   ", RTT=" & CStr(roundTripTime) & "ms", _
                   IIf(statusCode = 0, "IP_SUCCESS", "STATUS=" & CStr(statusCode)), _
                   ElapsedMilliseconds(startedAt)
    End If

CleanUp:
    If icmpHandle <> 0 And icmpHandle <> INVALID_HANDLE_VALUE Then
        closeResult = IcmpCloseHandle(icmpHandle)
        WriteTrace traceId, "CLIENT", "CLOSE_ICMP6_HANDLE", "LOCAL", _
                   "IcmpCloseHandle", _
                   IIf(closeResult <> 0, "OK", "ERROR=" & CStr(Err.LastDllError)), _
                   ElapsedMilliseconds(startedAt)
        icmpHandle = 0
    End If

    WriteTrace traceId, "CLIENT", "SUMMARY", "LOCAL", _
               "Destination=::1", _
               IIf(succeeded, "COMPLETED", "COMPLETED_WITHOUT_SUCCESS"), _
               ElapsedMilliseconds(startedAt)
    Exit Sub

UnexpectedError:
    WriteTrace traceId, "CLIENT", "UNEXPECTED_ERROR", "LOCAL", _
               "VBA Error " & CStr(Err.Number) & ": " & Err.Description, _
               "ERROR", ElapsedMilliseconds(startedAt)
    Resume CleanUp

End Sub

' 2バイト整数を指定位置へ書き込む。CPUのリトルエンディアン表現を利用する。
Private Sub PutInteger(ByRef buffer() As Byte, ByVal offset As Long, ByVal value As Integer)

    CopyMemory buffer(offset), value, 2

End Sub

Private Function ReadLong(ByRef buffer() As Byte, ByVal offset As Long) As Long

    Dim value As Long
    CopyMemory value, buffer(offset), 4
    ReadLong = value

End Function

' 16バイトのIPv6アドレスを8組の16進数として表示する。
' 省略表記(::)への圧縮はせず、メモリ上の8組が見える形式にしている。
Private Function ReadIPv6Address( _
    ByRef buffer() As Byte, _
    ByVal addressOffset As Long) As String

    Dim groupIndex As Long
    Dim groupValue As Long
    Dim result As String

    For groupIndex = 0 To 7
        groupValue = CLng(buffer(addressOffset + groupIndex * 2)) * 256& + _
                     CLng(buffer(addressOffset + groupIndex * 2 + 1))

        If groupIndex > 0 Then result = result & ":"
        result = result & LCase$(Hex$(groupValue))
    Next groupIndex

    ReadIPv6Address = result

End Function
