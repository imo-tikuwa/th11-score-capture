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
```
git clone https://github.com/imo-tikuwa/th11-score-capture
cd th11-score-capture
pip install -r requirements.txt
python app.rb
```

## メモ
スペルカードについてPhotoshopで切り抜きとかした後、そのままでは画像がうまく読み込めない模様。  
「libpng warning: iCCP: known incorrect sRGB profile」  というエラーが出てた。  
ImageMagickでexif情報を削除するバッチ(sh\strip_sample_spell_cards.bat)を実行する必要あり。

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
