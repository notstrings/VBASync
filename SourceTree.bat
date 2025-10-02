@echo off
setlocal

rem SourceTreeの実行ファイルパス
set "sourcetreePath=%LocalAppData%\SourceTree\SourceTree.exe"

rem カレントフォルダの取得
set "currentFolder=%cd%"

rem コマンドの組み立てと実行
"%sourcetreePath%" -f "%currentFolder%" status

endlocal