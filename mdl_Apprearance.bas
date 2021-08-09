Attribute VB_Name = "mdl_Apprearance"
'# アプリケーションの表示

''## リボンの表示
''- 表示
'   Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
''- 非表示
'   Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"

''## ステータスバーの表示
''- 表示
'   Application.DisplayStatusBar = True
''- 非表示
'   Application.DisplayStatusBar = False

''## 数式バーの表示
''- 表示
'    Application.DisplayFormulaBar = True
''- 非表示
'    Application.DisplayFormulaBar = False

''## タブの表示
''- 表示
'    ActiveWindow.DisplayHeadings = True
''- 非表示
'    ActiveWindow.DisplayHeadings = False

''## 画面を最前面に表示する
'    AppActivate Application.Caption

''## 画面のサイズ
''- 最大化する
'    Application.WindowState = xlMaximized
'- 最小化する
'    Application.WindowState = xlMinimized
'- 標準の大きさにする
'    Application.WindowState = xlNormal

''## アプリケーションエラーの発出
''- 発出する
'    Application.DisplayAlerts = True
''- 発出しない
'    Application.DisplayAlerts = False

''## 画面更新設定
''- 更新する
'    Applicaiton.ScreenUpdating = True
''- 更新しない
'    Application.ScreenUpdating = False

''## 切り取り機能の設定
''- 有効化
'    Application.CutCopyMode = 2
''- 無効化
'    Application.CutCopyMode = 0

''## イベントの抑止
''- 抑止解除
'    Application.EnableEvents = False
''- 抑止
'    Application.EnableEvents = True

''## シートの表示/非表示
''- 表示
'    Sheets("Sheet1").Visible = xlVisible
''- 非表示（シートタブ上で再表示設定可能）
'    Sheets("Sheet1").Visible = xlHidden
''- 完全に非表示（VBEのプロパティにて表示モードに変更可能）
'    Sheets("Sheet1").Visible = xlVeryHidden

''## シートの保護
''- 保護
'    Sheets("Sheet1").Protect
''- 保護解除
'    Sheets("Sheet1").Unprotect

''## ショートカットキーの設定
''- 無効化
'    Application.OnKey "%{F11}", ""
'    Application.OnKey "^x", ""
'    Application.OnKey "^s", ""
''- 有効化
'    Application.OnKey "%{F11}"
'    Application.OnKey "^x"
'    Application.OnKey "^s"
''- ショートカットキーの役割変更
'    Application.OnKey "^x", "^s"


