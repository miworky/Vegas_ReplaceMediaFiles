# Vegas_ReplaceMediaFiles
●概要
このプログラムは、Vegas Pro 18.0 のプラグインです。  
このプラグインを使用すると、タイムラインに配置してある「みてねからダウンロードしたメディアファイル」をオリジナルの高解像度のファイルに置き換えます。 

Take1に「みてねからダウンロードしたファイル」を張り付けておき、オリジナルの高解像度のファイルをメディアプールに登録しておいた上で本プラグインを実行すると、
Take2に「オリジナルの高解像度のファイル」が登録されます。

ファイルのマッチングは撮影日時で行います。

置き換え可能な形式

MP4, MOV, JPG

まだ置き換えできない形式（いずれ対応したい）

HEIF


●背景

みてねから画像や動画をダウンロードして動画を作成したり、ブルーレイに焼いたりする際に、いくつか課題があります。
課題の一つに、みてねにアップロードした画像や動画は、以下のように低解像度に変換されてしまっていることがあります：

画像：1920 x 1440

動画： 960 x 540 30fps

せっかくオリジナルの高解像度のファイルが手元にあるのに、低い解像度で動画を作るのは嫌なので、高解像度のファイルに差し替えたいのです。

●動画作成手順

1)みてねからファイルをダウンロードする

  https://github.com/miworky/miteneDownloader
  を使ってダウンロードしてください。
  ダウンロードしたファイルは、「YYYY-MM-DDThhmmss_1つめのコメント」というファイル名になります。
  
2)ダウンロードしたファイルを Vegas Pro のメディアプールに取り込み、タイムラインに貼り付けます。

　ダウンロードしたファイル名に日付が含まれているので、これだけでみてねからダウンロードしたファイルを撮影日時順にタイムラインに配置できます。

3)撮影日とコメントをテキストイベントに登録します

　後日公開予定のプラグインを使用すると自動でテキストイベントを追加できます。

4)オリジナルの高解像度のファイルに差し替えます（本プログラムを使用します）

5)お好きなBGMを貼り付けます

6)動画として書き出したり、ブルーレイに焼いたりします。


●開発環境

VisualStudio 2019 C#

●ビルド方法

1)ReplaceMediaFiles.sln を VisualStudio2019 で開きます

2)Release, Any CPUでビルド

ReplaceMediaFiles\bin\Releaseに成果物ができあがります。


●デプロイ方法

C:\ProgramData\VEGAS Pro\Script Menu
に以下のファイルをコピーします：

ReplaceMediaFiles\bin\Release\ReplaceMediaFiles.dll

ReplaceMediaFiles\bin\Release\MediaInfo.Wrapper.dll

ReplaceMediaFiles\bin\Release\x64 のファイルすべて


●実行方法

1)Vegas Pro 18.0 で、あらかじめ「みてねからダウンロードしたファイル」をタイムラインに配置して、オリジナルのファイルをメディアプールに登録しておきます

2)Vegas Pro 18.0 から本プラグインを実行します

3)ファイル選択ダイアログが開くので、作成するログのファイル名を指定します

4)しばらく時間が経った後に、「終了しました」というポップアップが開けば終了です

　2で選択したファイルに作業ログが出力されています。
 
 　作業ログには以下の内容が出力されます：
  
　　オリジナルのファイルが見つかった場合：元のファイル名とオリジナルのファイル名
  
　　オリジナルのファイルが見つからなかった場合：元のファイル名



