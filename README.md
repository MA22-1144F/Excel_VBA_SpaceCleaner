# Excel_VBA_SpaceCleaner

## 概要
Excelの選択範囲に含まれるスペースを一括削除します。  
PERSONAL.XLSBに保存して動作させることを想定しています。  
<img src="image_delete.png" alt="イメージ画像" width="700">

## 動作環境
Microsoft Excel上で動作します。  

## インストール方法
1. Contentsフォルダ内のUserForm削除選択.frm、UserForm削除選択.frx、拡張削除.basを任意の同一フォルダに保存
2. Excelで新規WorkSheetを開く
3. 開発タブのVisual BasicまたはAlt+F11でVBE(Visual Basic Editor)を開く
4. VBAProject一覧からVBAProject (PERSONAL.XLSB)を選択し右クリック
5. ファイルのインポートで保存したUserForm削除選択.frmを開く
6. 加えてファイルのインポートで保存した拡張削除.basを開く
7. 上書き保存ボタンまたはCtrl+Sで保存
以下オプション設定  
8. 任意のExcel WorkSheetに戻り、ファイル→オプション→リボンのユーザー設定の順で遷移
9. 任意のユーザー設定グループ(無ければ新規作成)にPERSONAL.XLSB!拡張削除.拡張削除のマクロを追加
10. お好みで名前やアイコンを変更
11. Excel WorkSheetに戻り、設定したマクロのアイコンを選択して起動

## 機能
* 選択範囲について以下のアクション  
  * 空白セルのみを削除して左方向にシフト  
  * 空白セルのみを削除して上方向にシフト  
  * 空白セルのある行全体を削除  
  * 空白セルのある列全体を削除  

## 連絡先
[Instagram](https://www.instagram.com/nattotoasto?igsh=NWNtdHhnY3A4NDQ0 "nattotoasto")

## ライセンス
MIT License
