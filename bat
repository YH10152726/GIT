@echo off
setlocal

:: =================================================================
:: Moldflow解析結果 XML一括出力バッチ for Excel評価ツール (複合ルール対応版)
:: 使用方法:
:: 1. このバッチファイルをMoldflowのスタディファイル(.sdy)と同じフォルダに置くか、
::    .sdyファイルをこのバッチファイルにドラッグ＆ドロップしてください。
:: 2. 実行すると「XML_Export」フォルダが作成され、その中にXMLファイルが出力されます。
:: =================================================================

:: --- 初期設定 ---
echo #################################################################
echo # Moldflow XML Exporter (Complex Rule Version)
echo #################################################################
echo.
if "%~1"=="" (
    echo ERROR: スタディファイル(.sdy)をこのバッチファイルにドラッグ＆ドロップしてください。
    echo.
    pause
    exit /b
)

set "SDY_FILE=%~1"
set "STUDY_NAME=%~n1"
set "OUTPUT_DIR=%~dp1XML_Export"

if not exist "%SDY_FILE%" (
    echo ERROR: 指定されたファイルが見つかりません。
    echo %SDY_FILE%
    echo.
    pause
    exit /b
)

echo 対象スタディファイル: %SDY_FILE%
echo 出力先フォルダ: %OUTPUT_DIR%
echo.
if not exist "%OUTPUT_DIR%" (
    echo "%OUTPUT_DIR%" を作成します...
    mkdir "%OUTPUT_DIR%"
)

:: --- studyrltコマンドのパス設定 ---
:: 環境変数にパスが通っていない場合、以下のパスをご自身の環境に合わせて変更してください。
set "STUDYRLT_PATH=studyrlt.exe"
:: 例: set "STUDYRLT_PATH=C:\Program Files\Autodesk\Moldflow Insight 2024\bin\studyrlt.exe"


:: --- XML出力処理 ---
echo --- XMLエクスポートを開始します ---

:: studyrltは出力ファイル名を指定できないため、一度デフォルト名で出力し、リネームします。
set "DEFAULT_XML_OUTPUT=%~dp1%STUDY_NAME%.xml"

:: --- 充填+保圧 (Fill+Pack) 結果 ---
call :ExportResult 1660 "confidence_of_fill" "充填の信頼性"
call :ExportResult 12220 "quality_prediction" "品質予測"
call :ExportResult 1540 "flow_front_temperature" "フローフロント温度"
call :ExportResult 10000 "pressure" "圧力"
call :ExportResult 4320 "shear_stress_at_wall" "壁面せん断応力"
call :ExportResult 4330 "viscosity" "粘度"
call :ExportResult 12330 "sink_marks_index" "ヒケ マーク インデックス"
call :ExportResult 1440 "total_part_weight" "部品重量 (XYプロット)"
call :ExportResult 8000 "clamp_force_xy" "型締力 (XYプロット)"
call :ExportResult 7000 "injection_pressure_xy" "射出圧力 (XYプロット)"
call :ExportResult 6230 "volumetric_shrinkage" "体積収縮"
call :ExportResult 1620 "frozen_layer_fraction" "スキン層の比率"
call :ExportResult 5010 "residual_stress" "残留応力"
call :ExportResult 4060 "orientation_index" "繊維配向"
call :ExportResult 12430 "birefringence" "複屈折"
call :ExportResult 10100 "pressure_at_vp_switch" "V/P切替時の圧力"

:: ★★★ ここから新規追加 (複合ルール用) ★★★
call :ExportResult 1640 "weld_lines" "ウェルドライン"
call :ExportResult 1670 "air_traps" "エアトラップ"
call :ExportResult 4310 "shear_rate_wall" "壁面せん断速度"
call :ExportResult 9260 "time_to_reach_ejection_temperature" "突き出し可能時間"

:: --- 冷却 (Cool) 結果 ---
call :ExportResult 9340 "cooling_time_variance" "冷却時間のばらつき"
call :ExportResult 9280 "time_to_freeze" "凍結時間"

:: --- そり (Warp) 結果 ---
call :ExportResult 13220 "deflection_total" "変位(全体)"
call :ExportResult 13230 "deflection_x" "変位(X)"
call :ExportResult 13240 "deflection_y" "変位(Y)"
call :ExportResult 13250 "deflection_z" "変位(Z)"


echo.
echo --- 全ての処理が完了しました ---
echo.
pause
goto :eof


:: --- 結果をエクスポートしてリネームするサブルーチン ---
:ExportResult
    set "RESULT_ID=%1"
    set "FILE_KEYWORD=%2"
    set "RESULT_NAME=%3"

    echo.
    echo %RESULT_NAME% (ID: %RESULT_ID%) をエクスポート中...
    "%STUDYRLT_PATH%" "%SDY_FILE%" -xml %RESULT_ID%
    
    if exist "%DEFAULT_XML_OUTPUT%" (
        echo   -> %FILE_KEYWORD%.xml にリネームします。
        move /Y "%DEFAULT_XML_OUTPUT%" "%OUTPUT_DIR%\%FILE_KEYWORD%.xml" > nul
    ) else (
        echo   WARN: XMLファイルの出力に失敗しました。結果IDが正しいか確認してください。
    )
    goto :eof
