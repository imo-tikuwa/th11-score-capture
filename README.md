# th11-score-capture

## このプログラムについて
東方地霊殿のウィンドウをリアルタイムでキャプチャして、スコア、グレイズ、スペルカード等の情報を
CSVファイルで出力するプログラムです。

## 開発環境について
|| バージョン等 |
|---|---|
| OS | Windows 10 Pro |
| Python | 3.8.2 |
| pip | 19.2.3 |

## インストール、起動
venvを利用
```
git clone https://github.com/imo-tikuwa/th11-score-capture
cd th11-score-capture
.\venv\Scripts\activate.bat
pip install -r requirements.txt
python app.rb
```

## オプション
| オプション名 | 内容 |
|---|---|
| -dev, --development | 開発モード(outputディレクトリに解析に使用した画像を保存) |
| --output | 指定したときコンソールに暫定の解析結果の出力を行います |
| --capture-period | 画面のキャプチャ間隔を指定してください(秒)、最低0～最大10.0 |
| --print-exec-time | 1回あたりのテンプレートマッチング処理全体の処理時間を出力します |

## 使い方
1. プログラムを実行する
```
python app.rb --output --capture-period 0.3
```

2.  (初回のみ)th11.exeの場所を指定する
3. プレイ画面を認識すると自動でスコア等のキャプチャおよび解析が開始します
4. 解析を終了したくなったところでCtrl+Cでプログラムの処理を中断
5. プログラムが終了します。解析した結果はoutputディレクトリに日付文字列付きのCSVファイルとして出力されます

## サンプルデータについて
sample_data以下のファイルを更新した場合はnpzファイルを更新する必要がある。  
具体的には直下にあるgenerate_npz_data.pyを実行するだけ。  
※サンプルデータのファイル自体を追加とかする場合はgenerate_npz_data.pyのrangeとかメンテナンスする必要あり。

## 履歴
2020/03/22  
ざっくり完成した模様  
めちゃくちゃ汚いコードなので直せるところは直したい

2020/03/23  
サンプルデータについてnpz化  
サンプルデータ周りは少し綺麗になったが他の部分はさらにコードが汚くなった
