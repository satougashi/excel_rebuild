めんせき
	・あくまで当方の問題解決のために作られたアドホック実装なので悪しからず
	・改造して使っていただいてかまいません

うさげ
	1. Rubyのインストールされたパソコンを準備する
	2. gem install rubyzipする
	3. githubから例のツールを取得する
	4. targetフォルダにxlsxファイルをzipにして解凍したファイルの中身を配置する
	5. $ ruby　./src/rebuild_excel_book.rbする
	6. resultフォルダに生成されたExcelファイルを上から叩いて、イカれた要素を特定する
