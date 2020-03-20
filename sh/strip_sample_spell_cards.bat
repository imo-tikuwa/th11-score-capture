@echo off
REM スペルカードのPNGはPhotoshopで左右を切り取る加工をしてるので、加工後のEXIF情報を削除する
REM 削除しないとOpenCvで画像を読み込むときに「libpng warning: iCCP: known incorrect sRGB profile」というエラーが出る模様

cd /D %~dp0
cd ..\sample_data\spell_card

REM ImageMagickでexif情報を削除
mogrify -strip *.png

REM 元のディレクトリに戻る
cd /D %~dp0