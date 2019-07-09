# AI Speaker
AI Speaker Client Server Sample Program


●サーバー側

WindowsのC#で実装
192.168.11.53のポート7000で待ち受ける

open wordという文字列を受け取ってWordを起動

close wordという文字列を受け取ってWordを保存終了

それ以外の文字列の場合、Wordを起動中なら受け取った文字列を書き込む


※Windowsファイアウォールを停止させて実行すること


●クラアント側

Raspberry Pi

Respeaker 2-Mics Pi HAT

Julius

を想定

pythonで実装

しゃべった単語を抽出して、192.168.11.53のポート7000に送る



