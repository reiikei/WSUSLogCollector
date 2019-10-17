# WSUSLogCollector について
日本マイクロソフトの WSUS サポート チームが作成した調査用のログ取得バッチファイルです。本バッチ ファイルを調査対象の WSUS サーバー上で管理者権限で実行することで、WSUS サーバー関連の情報を一括して採取することが可能です。事象の内容にもよりますが、通常採取した情報の解析には 3 営業日程度時間をいただいております。また、調査の状況によっては、調査の途中にて本バッチファイルでの情報採取以外にも、追加での情報採取を依頼する可能性があるためご留意ください。

# WSUSLogCollector の利用について
## WSUSLogCollector の実行要件
* WSUS がインストールされている Windows Server 2008 R2 以降の環境
* Powershell が利用可能な環境 (Windows Server 2008 R2 以降の環境であれば既定で利用可能)

## WSUSLogCollector の実行方法
1. [ZIP ファイル](https://github.com/reiikei/WSUSLogCollector/archive/1.0.0.zip) をダウンロード、展開し、WSUSLogCollector.cmd を情報採取対象の WSUS サーバーへコピーします。
2. 情報採取対象の WSUS サーバー上にて管理者権限を持つユーザーにて、WSUSLogCollector.cmd を右クリックし "管理者として実行" を選択して、情報の取得を開始します。
3. システム ドライブ (C ドライブ) 直下にファイル `WSUSLogs-<ホスト名>-<yyyymmddhhmnss>.zip` が出力されるため、本ファイルをお問い合わせの中でご案内するアップロード サイトより、アップロードしてください。

# よくあるご質問
<dl>
    <dt>Q. WSUSLogCollector による情報採取を行う際に、実行環境へ影響はありますか？</dt>
    <dd>WSUSLogCollector を実行することで、WSUS サーバーへ設定変更が行われることはありません。また、WSUSLogCollector を実行することで、実運用に影響を与えるレベルで実行環境に負荷を与えることは想定されませんが、ログ ファイルの出力やコピーに伴うリソースは利用します。パフォーマンスがシビアな環境で実施するためには、念のため、業務時間外等の負荷が低い時間帯での実行をご検討ください。</dd>
    <dt>Q. WSUSLogCollector にて出力されるログ ファイルは、どの程度の容量となりますか？</dt>
    <dd>環境によって異なるため、一概に申し上げることは出来ませんが、通常数十 MB 程度の容量となります。</dd>
    <dt>Q. WSUSLogCollector による情報の採取は、どのくらいの時間がかかりますか？</dt>
    <dd>環境によって異なるため、一概に申し上げることは出来ませんが、通常数分程度で完了します。</dd>
    <dt>Q. WSUSLogCollector によって、どのような情報が採取されますか？</dt>
    <dd>「WSUSLogCollector にて取得される情報の一覧」として後述しますので、こちらをご参考としてください。</dd>
</dl> 

# WSUSLogCollector にて取得出来る情報の一覧
WSUSLogCollector にて取得出来る情報の一覧について案内します。

## イベント ログ
* Application
* Security
* Setup
* System
* Microsoft-Windows-Bits-Client/Operational

## システム関連の情報
* Microsoft システム情報 (Msinfo32) の実行結果
* グループ ポリシーの結果セット
* OS にインストールされている更新プログラムの一覧 
* OS にインストールされている役割と機能の一覧 

## ネットワークおよび証明書関連情報
* 実行中の BITS ダウンロード ジョブの一覧
* ipconfig /all の出力
* WinHTTP のプロキシ設定
* hosts の設定
* 登録されているルート証明書の一覧

## WSUS 関連情報
* WSUS 関連のレジストリ情報 (`HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Update Services` の内容)
* WSUS のログ ファイル (`%ProgramFiles%\Update Services\LogFiles\` 配下のファイル)
* WSUS API の以下関数の結果 
    * [AdminProxy.GetUpdateServer Method ()](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms745830(v%3dvs.85))
    * [IUpdateServer.GetStatus Method ()](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms747050(v%3Dvs.85))
    * [IUpdateServer.GetConfiguration Method ()](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms747026(v=vs.85))
    * [IUpdateServer.GetSubscription Method ()](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms747052(v=vs.85))
    * [IUpdateServer.GetEmailNotificationConfiguration Method ()](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/aa349873(v=vs.85))
    * [IUpdateServer.GetDownstreamServers Method ()](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms747034(v=vs.85))
    * [IUpdateServer.GetDatabaseConfiguration Method ()](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms747031(v=vs.85))
* WSUSContent フォルダ配下のファイル一覧

## IIS 関連情報
* IIS の構成情報 (applicationHost.config および 各種 web.config)
* 一週間分の HTTPERR ログ
* 3 日分の IIS アクセス ログ (出力が有効にされている場合)

## データベース関連情報
* WID および Microsoft SQL Server のエラー ログ
