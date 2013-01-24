# lu2review

Liunx Users' Groupの[小江戸らぐ](http://koedolug.org/)が発行する活動報告集「Linux Users」の原稿（.odt形式）を、[ReVIEW](https://github.com/kmuto/review)形式に変換する。

## 使い方

    $ perl lu2review.pl foo.odt

カレントディレクトリにfoo.reが生成される。.odtファイルに画像が含まれていれば、カレントディレクトリにimagesディレクトリが作られ、その中に画像が保存される。
