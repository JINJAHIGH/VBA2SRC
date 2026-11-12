# VBA2SRC — Word・Access対応

Excelと同じ起動方法・保存先・UTF-8/LF・日本語JSONスキーマ2を使用します。
実行時にOffice、COM、ODBC、Javaは不要です。入力は一度だけ読み、以降はメモリ上で解析します。

| 対象 | 取得する保存情報 | Excel専用項目 |
| --- | --- | --- |
| Word OOXML（.docm/.dotm、VBAなしの.docx/.dotx） | VBA、Attribute、種類、プロジェクト、コンパイル引数、参照、保護 | 対象外 |
| Word 97以降のOLE（.doc/.dot） | 同上。WordDocumentのFIBと暗号化を先に検査 | 対象外 |
| Access Jet 4/ACE（.accdb、Access 2002以降の構造の.mdb） | 同上。現行のMSysAccessStorageにあるVBAProjectを取得 | 対象外 |

拡張子ではなく保存内容で識別します。Wordの主文書・VBAパーツはContent TypesとRelationshipsから解決し、固定のword/vbaProject.binには依存しません。
Accessはカタログから現行ストレージを読み、削除済みページやscratchからコードを探索しません。
長値（Lv）の復元はストリームごとに行い、一つのモジュールが壊れていても残りと設定を回収します。
Access解析にpyOpenVBA 6.6.0（MIT）を追加しています。内部RawRow APIへの依存はaccess.pyへ限定し、更新時は追加テストとEXE検証を実行してください。

WordのThisDocument、Accessフォーム／レポートのコードは、PROJECTのDocument宣言に従ってドキュメントモジュールへ保存します。
`--group-by-type` を指定すると、これらのコードは結果フォルダ内の `VBAまとめ/ドキュメントモジュール.txt` にも収録します。
名前がForm_やReport_から始まるだけでは分類しません。標準・クラスモジュールもPROJECTとdirで判定します。
コンパイル引数は保存原文と解析値、参照設定は保存されたLIBID・GUID・版・パス・順序を保持します。
参照の現在PC上での解決や、VBA・Accessマクロの実行はしません。dirに存在しない暗黙のホスト参照は合成しません。

VBAプロジェクト保護と文書／DB全体の暗号化は別です。前者は既存の候補照合設定を使用し、後者は復号が必要な失敗として返します。
破損・欠損・未対応保存構造はVBAなしへ置き換えません。正常なVBAなしは終了コード0です。

Accessフォーム／レポートのデザイン、テーブル定義・レコード・クエリ・Accessマクロは対象外です。
Jet 3、Access 2000のMSysAccessObjects構造、ソースのない.mde/.accdeの復元は未対応です。
データベースに保存されたソースを抽出する機能で、コンパイル済みp-codeの逆コンパイルは行いません。

検証には独自のWordパッケージと、依存ライブラリの空DBテンプレートへ独自VBAを格納したAccessデータを使用します。
条件付きコンパイル引数、参照、保護候補、日本語、標準／クラス／ドキュメント／フォーム、VBAなし、破損・暗号化・未対応、決定的出力、入力の不変性を確認します。
手元の保護付きテスト.docm／テスト.accdbでもソース・保存参照・プロジェクト情報の抽出を確認しました。これらはテストや配布ZIPへ同梱しません。

参照資料：[MS-OVBA](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/)
／[pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA)
／[MDB ToolsのJet/ACE構造資料](https://github.com/mdbtools/mdbtools/blob/dev/HACKING.md)。
