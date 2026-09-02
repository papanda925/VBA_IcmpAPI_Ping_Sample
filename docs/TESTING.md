# 動作確認手順

## 前提

- Windows 10／11
- VBA7を搭載したWindows版Excel（32bitまたは64bit）
- マクロ有効ブック（`.xlsm`）

このコードを作成した環境にはWindows版Excelがないため、以下は利用者のExcelで確認するための手順です。実機確認が終わるまでは、確認済みと扱わないでください。

## 1. インポートとコンパイル

1. 新しい`.xlsm`ブックを作成する。
2. `Alt`+`F11`でVBEを開く。
3. `TraceLogger.bas`、`IcmpPingIPv4.bas`、`IcmpPingIPv6Loopback.bas`をインポートする。
4. 「デバッグ」→「VBAProjectのコンパイル」を押す。
5. コンパイルエラーが表示されないことを確認する。

エラーが出た場合は、Officeの製品名、32／64bit、エラー行、エラーメッセージを記録してください。内部IPアドレス等は公開Issueへ載せないでください。

## 2. IPv4ループバック

`RunPingLocalhost`を実行します。

期待する結果：

- `Trace`シートが作成される。
- `OPEN_ICMP_HANDLE`が`OK`になる。
- 4つの`SEND_ECHO`と、それぞれに対応する`RECEIVE_REPLY`が記録される。
- 通常は`From=127.0.0.1`、`IP_SUCCESS`となる。
- 最後に`CLOSE_ICMP_HANDLE`と`SUMMARY`が記録される。

## 3. IPv6ループバック

`RunPingIPv6Loopback`を実行します。

期待する結果：

- `BUILD_ADDRESS`で28バイトの`SOCKADDR_IN6`が作成される。
- `OPEN_ICMP6_HANDLE`が`OK`になる。
- `RECEIVE_REPLY`が`IP_SUCCESS`になる。
- 応答元は省略しない形式の`0:0:0:0:0:0:0:1`で表示される。

IPv6スタックが無効な環境では、`ERROR_NOT_SUPPORTED`相当のエラーになる可能性があります。

## 4. 入力値検証

イミディエイトウィンドウで次を試します。

```vb
? PingIPv4("999.0.0.1", 1, 1000)
? PingIPv4("127.0.0.1", 0, 1000)
? PingIPv4("127.0.0.1", 11, 1000)
? PingIPv4("127.0.0.1", 1, 0)
? PingIPv4("127.0.0.1", 1, 60001)
```

いずれもAPI送信前にFalseとなり、Traceシートへ`INVALID_*`が記録されることを確認します。

## 5. タイムアウト

許可された試験用アドレスがある場合だけ、応答しないことが分かっているアドレスへ1回送信します。

```vb
? PingIPv4("192.0.2.1", 1, 500)
```

`192.0.2.0/24`は文書例示用ですが、組織の経路設定によって挙動は異なります。`IP_REQ_TIMED_OUT`等が記録されても、コード異常とは限りません。

## 6. 32bit／64bit確認票

| 確認項目 | 32bit Excel | 64bit Excel |
|---|---:|---:|
| VBAProjectのコンパイル | 未確認 | 未確認 |
| IPv4 `127.0.0.1` | 未確認 | 未確認 |
| IPv6 `::1` | 未確認 | 未確認 |
| ハンドル解放 | 未確認 | 未確認 |
| Traceシート | 未確認 | 未確認 |

確認後は、OS・Excelのバージョンとともに表を更新してください。
