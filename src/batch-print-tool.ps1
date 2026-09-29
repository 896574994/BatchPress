# ============================================================
#  Batch Print Tool - Local (PowerShell / WinForms)
#  Multilingual (简体中文 / English)
#  Features:
#    1. List and select all installed printers on this PC
#    2. Batch settings: copies, one/two-sided, orientation
#    3. Silent batch printing, no per-file confirmation
#  Supported: Images (PNG/JPG/BMP/GIF/TIFF), TXT, PDF (SumatraPDF),
#             Word (doc/docx via Office or WPS), Excel (xls/xlsx)
#  Usage: double-click start.bat, or
#         powershell -ExecutionPolicy Bypass -File this-script.ps1
# ============================================================

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.Application]::EnableVisualStyles()

# Win32 print API: switch the system default printer (more reliable than
# Office's ActivePrinter, which fails on some machines / printer ports)
if (-not ('PrintApi' -as [type])) {
    Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public class PrintApi {
    [DllImport("winspool.drv", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern bool SetDefaultPrinter(string p);
    [DllImport("winspool.drv", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern bool GetDefaultPrinter(System.Text.StringBuilder buf, ref int size);
}
'@
}
function Get-SystemDefaultPrinter {
    $sb = New-Object System.Text.StringBuilder 512
    $size = $sb.Capacity
    if ([PrintApi]::GetDefaultPrinter($sb, [ref]$size)) { return $sb.ToString() }
    return $null
}

$script:fileList = New-Object System.Collections.ArrayList   # files to print
$script:abort = $false

# ---------------- Language strings ----------------
$script:S = @{
    'zh-CN' = @{
        'formTitle' = '批量打印工具 - 本地版'
        'lblPrinter' = '打印机：'
        'lblLang' = '语言：'
        'grpParam' = '打印参数'
        'lblCopies' = '份数：'
        'lblDuplex' = '双面：'
        'lblOrient' = '方向：'
        'lblTip' = '提示：若列表中有打印机未列出，请先在系统设置中添加打印机后重新打开本工具'
        'btnAdd' = '添加文件'
        'btnRemove' = '移除选中'
        'btnClear' = '清空列表'
        'lblCount' = '共 {0} 个文件'
        'btnPrint' = '开始打印'
        'btnPrinting' = '打印中...'
        'dlgTitle' = '选择要批量打印的文件'
        'dlgFilter' = '支持的文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.txt;*.pdf;*.doc;*.docx;*.xls;*.xlsx|图片|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|文本|*.txt|PDF|*.pdf|Word|*.doc;*.docx|Excel|*.xls;*.xlsx|所有文件|*.*'
        'dupOne' = '单面'
        'dupLong' = '双面-长边翻转'
        'dupShort' = '双面-短边翻转'
        'orientAuto' = '自动'
        'orientPortrait' = '纵向'
        'orientLandscape' = '横向'
        'msgNoFiles' = '请先添加要打印的文件'
        'msgNoPrinter' = '未检测到打印机，请先在系统中添加打印机'
        'msgTitle' = '提示'
        'logTitle' = '批量打印工具 - 本地版'
        'logPrinterCount' = '检测到打印机：{0} 台'
        'logNoPrinter' = '[!!] 未检测到打印机，请先在系统设置中添加打印机'
        'logSumatraYes' = '检测到 SumatraPDF，支持 PDF 静默打印'
        'logSumatraNo' = '未检测到 SumatraPDF：PDF 文件将提示手动打印（安装后自动支持）'
        'logWordMs' = '检测到 Microsoft Word，支持 Word 文档打印'
        'logWordWps' = '检测到 WPS 文字，支持 Word 文档打印'
        'logWordNo' = '[!!] 未检测到 Word/WPS，doc/docx 将无法打印'
        'logExcelMs' = '检测到 Microsoft Excel，支持表格打印'
        'logExcelWps' = '检测到 WPS 表格，支持表格打印'
        'logExcelNo' = '[!!] 未检测到 Excel/WPS 表格，xls/xlsx 将无法打印'
        'logPrinter' = '打印机：{0}'
        'logSettings' = '参数：份数={0}  双面={1}  方向={2}'
        'logStart' = '开始批量打印，共 {0} 个文件'
        'logAbort' = '用户中止打印'
        'logDuplexApplied' = '双面模式已应用：{0}'
        'logDuplexFail' = '[!!] 无法设置打印机双面模式（{0}），将按打印机现有设置打印'
        'logDuplexRestore' = '已恢复打印机原双面设置：{0}'
        'logDuplexRestoreFail' = '[!!] 恢复打印机双面设置失败：{0}'
        'logImage' = '[OK] 图片 {0}'
        'logText' = '[OK] 文本 {0}'
        'logPdf' = '[OK] PDF {0}'
        'logWord' = '[OK] Word {0}'
        'logExcel' = '[OK] Excel {0}'
        'logFail' = '[!!] 打印失败：{0}'
        'logSkip' = '[--] 跳过不支持的格式：{0} ({1})'
        'logDone' = '完成：成功 {0} 个，失败 {1} 个'
        'logPdfNeed' = 'PDF 需要 SumatraPDF：{0}'
        'logPdfTip' = '请安装 SumatraPDF（免费）后重试，或改用浏览器打开该 PDF 手动打印'
        'errNoExcel' = '本机未安装 Excel 或 WPS 表格，无法打印 {0} 文件'
        'errNoWord' = '本机未安装 Word 或 WPS 文字，无法打印 {0} 文件'
        'errSwitchPrinter' = '无法将打印机切换为 [{0}]，请确认打印机名称'
        'fileMissing' = '缺失/无法访问'
    }
    'en-US' = @{
        'formTitle' = 'Batch Print Tool - Local'
        'lblPrinter' = 'Printer: '
        'lblLang' = 'Language: '
        'grpParam' = 'Print Settings'
        'lblCopies' = 'Copies: '
        'lblDuplex' = 'Duplex: '
        'lblOrient' = 'Orientation: '
        'lblTip' = 'Tip: If your printer is not listed, add it in Windows Settings first, then restart this tool.'
        'btnAdd' = 'Add Files'
        'btnRemove' = 'Remove Selected'
        'btnClear' = 'Clear List'
        'lblCount' = '{0} file(s)'
        'btnPrint' = 'Start Printing'
        'btnPrinting' = 'Printing...'
        'dlgTitle' = 'Select files to batch print'
        'dlgFilter' = 'Supported files|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.txt;*.pdf;*.doc;*.docx;*.xls;*.xlsx|Images|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|Text|*.txt|PDF|*.pdf|Word|*.doc;*.docx|Excel|*.xls;*.xlsx|All files|*.*'
        'dupOne' = 'One-sided'
        'dupLong' = 'Duplex - Long Edge'
        'dupShort' = 'Duplex - Short Edge'
        'orientAuto' = 'Auto'
        'orientPortrait' = 'Portrait'
        'orientLandscape' = 'Landscape'
        'msgNoFiles' = 'Please add files to print first'
        'msgNoPrinter' = 'No printer detected. Please add a printer in Windows Settings first.'
        'msgTitle' = 'Notice'
        'logTitle' = 'Batch Print Tool - Local'
        'logPrinterCount' = 'Printers detected: {0}'
        'logNoPrinter' = '[!!] No printer detected. Please add a printer in Windows Settings first.'
        'logSumatraYes' = 'SumatraPDF detected, silent PDF printing supported'
        'logSumatraNo' = 'SumatraPDF not found: PDF files will be printed manually (auto-supported after install)'
        'logWordMs' = 'Microsoft Word detected, Word documents supported'
        'logWordWps' = 'WPS Writer detected, Word documents supported'
        'logWordNo' = '[!!] Word/WPS not found, doc/docx cannot be printed'
        'logExcelMs' = 'Microsoft Excel detected, spreadsheets supported'
        'logExcelWps' = 'WPS Spreadsheets detected, spreadsheets supported'
        'logExcelNo' = '[!!] Excel/WPS not found, xls/xlsx cannot be printed'
        'logPrinter' = 'Printer: {0}'
        'logSettings' = 'Settings: Copies={0}  Duplex={1}  Orientation={2}'
        'logStart' = 'Starting batch print, {0} file(s)'
        'logAbort' = 'Print aborted by user'
        'logDuplexApplied' = 'Duplex mode applied: {0}'
        'logDuplexFail' = '[!!] Failed to set printer duplex mode ({0}), printing with current settings'
        'logDuplexRestore' = 'Restored printer duplex setting: {0}'
        'logDuplexRestoreFail' = '[!!] Failed to restore printer duplex setting: {0}'
        'logImage' = '[OK] Image {0}'
        'logText' = '[OK] Text {0}'
        'logPdf' = '[OK] PDF {0}'
        'logWord' = '[OK] Word {0}'
        'logExcel' = '[OK] Excel {0}'
        'logFail' = '[!!] Print failed: {0}'
        'logSkip' = '[--] Skipped unsupported format: {0} ({1})'
        'logDone' = 'Done: {0} succeeded, {1} failed'
        'logPdfNeed' = 'PDF requires SumatraPDF: {0}'
        'logPdfTip' = 'Install SumatraPDF (free), or open the PDF in a browser and print manually.'
        'errNoExcel' = 'Excel or WPS Spreadsheets is not installed, cannot print {0} files'
        'errNoWord' = 'Word or WPS Writer is not installed, cannot print {0} files'
        'errSwitchPrinter' = 'Cannot switch to printer [{0}], please check the printer name'
        'fileMissing' = 'missing / inaccessible'
    }
    'zh-TW' = @{
        'formTitle' = '批次列印工具 - 本機版'
        'lblPrinter' = '印表機：'
        'lblLang' = '語言：'
        'grpParam' = '列印參數'
        'lblCopies' = '份數：'
        'lblDuplex' = '雙面：'
        'lblOrient' = '方向：'
        'lblTip' = '提示：若列表中有印表機未列出，請先在系統設定中新增印表機後重新開啟本工具'
        'btnAdd' = '新增檔案'
        'btnRemove' = '移除選取'
        'btnClear' = '清空清單'
        'lblCount' = '共 {0} 個檔案'
        'btnPrint' = '開始列印'
        'btnPrinting' = '列印中...'
        'dlgTitle' = '選擇要批次列印的檔案'
        'dlgFilter' = '支援的檔案|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.txt;*.pdf;*.doc;*.docx;*.xls;*.xlsx|圖片|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|文字|*.txt|PDF|*.pdf|Word|*.doc;*.docx|Excel|*.xls;*.xlsx|所有檔案|*.*'
        'dupOne' = '單面'
        'dupLong' = '雙面-長邊翻頁'
        'dupShort' = '雙面-短邊翻頁'
        'orientAuto' = '自動'
        'orientPortrait' = '直向'
        'orientLandscape' = '橫向'
        'msgNoFiles' = '請先新增要列印的檔案'
        'msgNoPrinter' = '未偵測到印表機，請先在系統中新增印表機'
        'msgTitle' = '提示'
        'logTitle' = '批次列印工具 - 本機版'
        'logPrinterCount' = '偵測到印表機：{0} 台'
        'logNoPrinter' = '[!!] 未偵測到印表機，請先在系統設定中新增印表機'
        'logSumatraYes' = '偵測到 SumatraPDF，支援 PDF 靜默列印'
        'logSumatraNo' = '未偵測到 SumatraPDF：PDF 檔案將提示手動列印（安裝後自動支援）'
        'logWordMs' = '偵測到 Microsoft Word，支援 Word 文件列印'
        'logWordWps' = '偵測到 WPS 文書，支援 Word 文件列印'
        'logWordNo' = '[!!] 未偵測到 Word/WPS，doc/docx 將無法列印'
        'logExcelMs' = '偵測到 Microsoft Excel，支援試算表列印'
        'logExcelWps' = '偵測到 WPS 試算表，支援試算表列印'
        'logExcelNo' = '[!!] 未偵測到 Excel/WPS 試算表，xls/xlsx 將無法列印'
        'logPrinter' = '印表機：{0}'
        'logSettings' = '參數：份數={0}  雙面={1}  方向={2}'
        'logStart' = '開始批次列印，共 {0} 個檔案'
        'logAbort' = '使用者中止列印'
        'logDuplexApplied' = '雙面模式已套用：{0}'
        'logDuplexFail' = '[!!] 無法設定印表機雙面模式（{0}），將依印表機現有設定列印'
        'logDuplexRestore' = '已還原印表機原雙面設定：{0}'
        'logDuplexRestoreFail' = '[!!] 還原印表機雙面設定失敗：{0}'
        'logImage' = '[OK] 圖片 {0}'
        'logText' = '[OK] 文字 {0}'
        'logPdf' = '[OK] PDF {0}'
        'logWord' = '[OK] Word {0}'
        'logExcel' = '[OK] Excel {0}'
        'logFail' = '[!!] 列印失敗：{0}'
        'logSkip' = '[--] 略過不支援的格式：{0} ({1})'
        'logDone' = '完成：成功 {0} 個，失敗 {1} 個'
        'logPdfNeed' = 'PDF 需要 SumatraPDF：{0}'
        'logPdfTip' = '請安裝 SumatraPDF（免費）後重試，或改用瀏覽器開啟該 PDF 手動列印'
        'errNoExcel' = '本機未安裝 Excel 或 WPS 試算表，無法列印 {0} 檔案'
        'errNoWord' = '本機未安裝 Word 或 WPS 文書，無法列印 {0} 檔案'
        'errSwitchPrinter' = '無法將印表機切換為 [{0}]，請確認印表機名稱'
        'fileMissing' = '遺失/無法存取'
    }
    'ja-JP' = @{
        'formTitle' = '一括印刷ツール - ローカル版'
        'lblPrinter' = 'プリンター：'
        'lblLang' = '言語：'
        'grpParam' = '印刷設定'
        'lblCopies' = '部数：'
        'lblDuplex' = '両面：'
        'lblOrient' = '向き：'
        'lblTip' = 'ヒント：リストにプリンターが表示されない場合は、まずシステム設定でプリンターを追加してから本ツールを再起動してください'
        'btnAdd' = 'ファイルを追加'
        'btnRemove' = '選択を削除'
        'btnClear' = 'リストをクリア'
        'lblCount' = '合計 {0} ファイル'
        'btnPrint' = '印刷開始'
        'btnPrinting' = '印刷中...'
        'dlgTitle' = '一括印刷するファイルを選択'
        'dlgFilter' = '対応ファイル|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.txt;*.pdf;*.doc;*.docx;*.xls;*.xlsx|画像|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|テキスト|*.txt|PDF|*.pdf|Word|*.doc;*.docx|Excel|*.xls;*.xlsx|すべてのファイル|*.*'
        'dupOne' = '片面'
        'dupLong' = '両面-長辺とじ'
        'dupShort' = '両面-短辺とじ'
        'orientAuto' = '自動'
        'orientPortrait' = '縦向き'
        'orientLandscape' = '横向き'
        'msgNoFiles' = '印刷するファイルを先に追加してください'
        'msgNoPrinter' = 'プリンターが検出されません。システムでプリンターを追加してください'
        'msgTitle' = 'お知らせ'
        'logTitle' = '一括印刷ツール - ローカル版'
        'logPrinterCount' = '検出したプリンター：{0} 台'
        'logNoPrinter' = '[!!] プリンターが検出されません。システム設定でプリンターを追加してください'
        'logSumatraYes' = 'SumatraPDF を検出しました。PDF のサイレント印刷に対応'
        'logSumatraNo' = 'SumatraPDF が見つかりません：PDF は手動印刷になります（インストール後は自動対応）'
        'logWordMs' = 'Microsoft Word を検出しました。Word 文書の印刷に対応'
        'logWordWps' = 'WPS 文字を検出しました。Word 文書の印刷に対応'
        'logWordNo' = '[!!] Word/WPS が見つかりません。doc/docx は印刷できません'
        'logExcelMs' = 'Microsoft Excel を検出しました。表計算の印刷に対応'
        'logExcelWps' = 'WPS 表計算を検出しました。表計算の印刷に対応'
        'logExcelNo' = '[!!] Excel/WPS 表計算が見つかりません。xls/xlsx は印刷できません'
        'logPrinter' = 'プリンター：{0}'
        'logSettings' = '設定：部数={0}  両面={1}  向き={2}'
        'logStart' = '一括印刷を開始します。合計 {0} ファイル'
        'logAbort' = 'ユーザーが印刷を中止しました'
        'logDuplexApplied' = '両面モードを適用：{0}'
        'logDuplexFail' = '[!!] プリンターの両面モードを設定できません（{0}）。現在の設定で印刷します'
        'logDuplexRestore' = 'プリンターの元の両面設定を復元：{0}'
        'logDuplexRestoreFail' = '[!!] プリンターの両面設定の復元に失敗：{0}'
        'logImage' = '[OK] 画像 {0}'
        'logText' = '[OK] テキスト {0}'
        'logPdf' = '[OK] PDF {0}'
        'logWord' = '[OK] Word {0}'
        'logExcel' = '[OK] Excel {0}'
        'logFail' = '[!!] 印刷失敗：{0}'
        'logSkip' = '[--] 対応しない形式をスキップ：{0} ({1})'
        'logDone' = '完了：成功 {0} 件、失敗 {1} 件'
        'logPdfNeed' = 'PDF には SumatraPDF が必要：{0}'
        'logPdfTip' = 'SumatraPDF（無料）をインストールして再試行するか、ブラウザーで PDF を開いて手動印刷してください'
        'errNoExcel' = 'Excel または WPS 表計算がインストールされていないため、{0} ファイルを印刷できません'
        'errNoWord' = 'Word または WPS 文字がインストールされていないため、{0} ファイルを印刷できません'
        'errSwitchPrinter' = 'プリンターを [{0}] に切り替えられません。プリンター名を確認してください'
        'fileMissing' = '見つからない/アクセス不可'
    }
    'ko-KR' = @{
        'formTitle' = '일괄 인쇄 도구 - 로컬 버전'
        'lblPrinter' = '프린터:'
        'lblLang' = '언어:'
        'grpParam' = '인쇄 설정'
        'lblCopies' = '매수:'
        'lblDuplex' = '양면:'
        'lblOrient' = '방향:'
        'lblTip' = '팁: 목록에 프린터가 없으면 시스템 설정에서 프린터를 먼저 추가한 후 이 도구를 다시 실행하세요'
        'btnAdd' = '파일 추가'
        'btnRemove' = '선택 삭제'
        'btnClear' = '목록 지우기'
        'lblCount' = '총 {0}개 파일'
        'btnPrint' = '인쇄 시작'
        'btnPrinting' = '인쇄 중...'
        'dlgTitle' = '일괄 인쇄할 파일 선택'
        'dlgFilter' = '지원 파일|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.txt;*.pdf;*.doc;*.docx;*.xls;*.xlsx|이미지|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|텍스트|*.txt|PDF|*.pdf|Word|*.doc;*.docx|Excel|*.xls;*.xlsx|모든 파일|*.*'
        'dupOne' = '단면'
        'dupLong' = '양면-긴 쪽 넘김'
        'dupShort' = '양면-짧은 쪽 넘김'
        'orientAuto' = '자동'
        'orientPortrait' = '세로'
        'orientLandscape' = '가로'
        'msgNoFiles' = '인쇄할 파일을 먼저 추가하세요'
        'msgNoPrinter' = '프린터가 감지되지 않았습니다. 시스템에서 프린터를 추가하세요'
        'msgTitle' = '알림'
        'logTitle' = '일괄 인쇄 도구 - 로컬 버전'
        'logPrinterCount' = '감지된 프린터: {0}대'
        'logNoPrinter' = '[!!] 프린터가 감지되지 않았습니다. 시스템 설정에서 프린터를 추가하세요'
        'logSumatraYes' = 'SumatraPDF가 감지되었습니다. PDF 자동 인쇄 지원'
        'logSumatraNo' = 'SumatraPDF를 찾을 수 없음: PDF 파일은 수동 인쇄됩니다 (설치 후 자동 지원)'
        'logWordMs' = 'Microsoft Word가 감지되었습니다. Word 문서 인쇄 지원'
        'logWordWps' = 'WPS 워드가 감지되었습니다. Word 문서 인쇄 지원'
        'logWordNo' = '[!!] Word/WPS를 찾을 수 없음, doc/docx를 인쇄할 수 없습니다'
        'logExcelMs' = 'Microsoft Excel이 감지되었습니다. 스프레드시트 인쇄 지원'
        'logExcelWps' = 'WPS 스프레드시트가 감지되었습니다. 스프레드시트 인쇄 지원'
        'logExcelNo' = '[!!] Excel/WPS 스프레드시트를 찾을 수 없음, xls/xlsx를 인쇄할 수 없습니다'
        'logPrinter' = '프린터: {0}'
        'logSettings' = '설정: 매수={0}  양면={1}  방향={2}'
        'logStart' = '일괄 인쇄 시작, 총 {0}개 파일'
        'logAbort' = '사용자가 인쇄를 중단했습니다'
        'logDuplexApplied' = '양면 모드 적용: {0}'
        'logDuplexFail' = '[!!] 프린터 양면 모드를 설정할 수 없음 ({0}), 현재 설정으로 인쇄합니다'
        'logDuplexRestore' = '프린터 원래 양면 설정 복원: {0}'
        'logDuplexRestoreFail' = '[!!] 프린터 양면 설정 복원 실패: {0}'
        'logImage' = '[OK] 이미지 {0}'
        'logText' = '[OK] 텍스트 {0}'
        'logPdf' = '[OK] PDF {0}'
        'logWord' = '[OK] Word {0}'
        'logExcel' = '[OK] Excel {0}'
        'logFail' = '[!!] 인쇄 실패: {0}'
        'logSkip' = '[--] 지원하지 않는 형식 건너뜀: {0} ({1})'
        'logDone' = '완료: 성공 {0}개, 실패 {1}개'
        'logPdfNeed' = 'PDF에는 SumatraPDF가 필요합니다: {0}'
        'logPdfTip' = 'SumatraPDF(무료)를 설치한 후 다시 시도하거나 브라우저에서 PDF를 열어 수동으로 인쇄하세요'
        'errNoExcel' = 'Excel 또는 WPS 스프레드시트가 설치되지 않아 {0} 파일을 인쇄할 수 없습니다'
        'errNoWord' = 'Word 또는 WPS 워드가 설치되지 않아 {0} 파일을 인쇄할 수 없습니다'
        'errSwitchPrinter' = '프린터를 [{0}]로 전환할 수 없습니다. 프린터 이름을 확인하세요'
        'fileMissing' = '없음/액세스 불가'
    }
    'de-DE' = @{
        'formTitle' = 'Stapeldruck-Tool - Lokalversion'
        'lblPrinter' = 'Drucker: '
        'lblLang' = 'Sprache: '
        'grpParam' = 'Druckeinstellungen'
        'lblCopies' = 'Exemplare: '
        'lblDuplex' = 'Duplex: '
        'lblOrient' = 'Ausrichtung: '
        'lblTip' = 'Tipp: Wenn Ihr Drucker nicht aufgelistet ist, fügen Sie ihn zuerst in den Windows-Einstellungen hinzu und starten Sie das Tool neu.'
        'btnAdd' = 'Dateien hinzufügen'
        'btnRemove' = 'Auswahl entfernen'
        'btnClear' = 'Liste leeren'
        'lblCount' = '{0} Datei(en)'
        'btnPrint' = 'Druck starten'
        'btnPrinting' = 'Drucke...'
        'dlgTitle' = 'Dateien für den Stapeldruck auswählen'
        'dlgFilter' = 'Unterstützte Dateien|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.txt;*.pdf;*.doc;*.docx;*.xls;*.xlsx|Bilder|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|Text|*.txt|PDF|*.pdf|Word|*.doc;*.docx|Excel|*.xls;*.xlsx|Alle Dateien|*.*'
        'dupOne' = 'Einseitig'
        'dupLong' = 'Duplex - lange Kante'
        'dupShort' = 'Duplex - kurze Kante'
        'orientAuto' = 'Automatisch'
        'orientPortrait' = 'Hochformat'
        'orientLandscape' = 'Querformat'
        'msgNoFiles' = 'Bitte zuerst Dateien zum Drucken hinzufügen'
        'msgNoPrinter' = 'Kein Drucker erkannt. Bitte zuerst einen Drucker in den Windows-Einstellungen hinzufügen.'
        'msgTitle' = 'Hinweis'
        'logTitle' = 'Stapeldruck-Tool - Lokalversion'
        'logPrinterCount' = 'Erkannte Drucker: {0}'
        'logNoPrinter' = '[!!] Kein Drucker erkannt. Bitte zuerst einen Drucker in den Windows-Einstellungen hinzufügen.'
        'logSumatraYes' = 'SumatraPDF erkannt, PDF-Stilldruck unterstützt'
        'logSumatraNo' = 'SumatraPDF nicht gefunden: PDF-Dateien werden manuell gedruckt (nach Installation automatisch)'
        'logWordMs' = 'Microsoft Word erkannt, Word-Dokumente unterstützt'
        'logWordWps' = 'WPS Writer erkannt, Word-Dokumente unterstützt'
        'logWordNo' = '[!!] Word/WPS nicht gefunden, doc/docx können nicht gedruckt werden'
        'logExcelMs' = 'Microsoft Excel erkannt, Tabellen unterstützt'
        'logExcelWps' = 'WPS Tabellen erkannt, Tabellen unterstützt'
        'logExcelNo' = '[!!] Excel/WPS Tabellen nicht gefunden, xls/xlsx können nicht gedruckt werden'
        'logPrinter' = 'Drucker: {0}'
        'logSettings' = 'Einstellungen: Exemplare={0}  Duplex={1}  Ausrichtung={2}'
        'logStart' = 'Stapeldruck gestartet, {0} Datei(en)'
        'logAbort' = 'Druck vom Benutzer abgebrochen'
        'logDuplexApplied' = 'Duplexmodus angewendet: {0}'
        'logDuplexFail' = '[!!] Duplexmodus des Druckers konnte nicht gesetzt werden ({0}), Druck mit aktuellen Einstellungen'
        'logDuplexRestore' = 'Duplex-Einstellung des Druckers wiederhergestellt: {0}'
        'logDuplexRestoreFail' = '[!!] Duplex-Einstellung des Druckers konnte nicht wiederhergestellt werden: {0}'
        'logImage' = '[OK] Bild {0}'
        'logText' = '[OK] Text {0}'
        'logPdf' = '[OK] PDF {0}'
        'logWord' = '[OK] Word {0}'
        'logExcel' = '[OK] Excel {0}'
        'logFail' = '[!!] Druck fehlgeschlagen: {0}'
        'logSkip' = '[--] Nicht unterstütztes Format übersprungen: {0} ({1})'
        'logDone' = 'Fertig: {0} erfolgreich, {1} fehlgeschlagen'
        'logPdfNeed' = 'PDF benötigt SumatraPDF: {0}'
        'logPdfTip' = 'SumatraPDF (kostenlos) installieren und erneut versuchen, oder PDF im Browser öffnen und manuell drucken.'
        'errNoExcel' = 'Excel oder WPS Tabellen ist nicht installiert, {0} Dateien können nicht gedruckt werden'
        'errNoWord' = 'Word oder WPS Writer ist nicht installiert, {0} Dateien können nicht gedruckt werden'
        'errSwitchPrinter' = 'Kann nicht zu Drucker [{0}] wechseln, bitte Druckername prüfen'
        'fileMissing' = 'fehlt / nicht erreichbar'
    }
    'fr-FR' = @{
        'formTitle' = 'Outil d''impression par lot - Version locale'
        'lblPrinter' = 'Imprimante : '
        'lblLang' = 'Langue : '
        'grpParam' = 'Paramètres d''impression'
        'lblCopies' = 'Copies : '
        'lblDuplex' = 'Recto-verso : '
        'lblOrient' = 'Orientation : '
        'lblTip' = 'Astuce : si votre imprimante n''est pas listée, ajoutez-la d''abord dans les paramètres Windows, puis redémarrez cet outil.'
        'btnAdd' = 'Ajouter des fichiers'
        'btnRemove' = 'Retirer la sélection'
        'btnClear' = 'Vider la liste'
        'lblCount' = '{0} fichier(s)'
        'btnPrint' = 'Imprimer'
        'btnPrinting' = 'Impression...'
        'dlgTitle' = 'Sélectionner les fichiers à imprimer en lot'
        'dlgFilter' = 'Fichiers pris en charge|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.txt;*.pdf;*.doc;*.docx;*.xls;*.xlsx|Images|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|Texte|*.txt|PDF|*.pdf|Word|*.doc;*.docx|Excel|*.xls;*.xlsx|Tous les fichiers|*.*'
        'dupOne' = 'Recto seul'
        'dupLong' = 'Recto-verso - bord long'
        'dupShort' = 'Recto-verso - bord court'
        'orientAuto' = 'Automatique'
        'orientPortrait' = 'Portrait'
        'orientLandscape' = 'Paysage'
        'msgNoFiles' = 'Veuillez d''abord ajouter des fichiers à imprimer'
        'msgNoPrinter' = 'Aucune imprimante détectée. Veuillez d''abord ajouter une imprimante dans les paramètres Windows.'
        'msgTitle' = 'Avis'
        'logTitle' = 'Outil d''impression par lot - Version locale'
        'logPrinterCount' = 'Imprimantes détectées : {0}'
        'logNoPrinter' = '[!!] Aucune imprimante détectée. Veuillez d''abord ajouter une imprimante dans les paramètres Windows.'
        'logSumatraYes' = 'SumatraPDF détecté, impression PDF silencieuse prise en charge'
        'logSumatraNo' = 'SumatraPDF introuvable : les PDF seront imprimés manuellement (pris en charge après installation)'
        'logWordMs' = 'Microsoft Word détecté, documents Word pris en charge'
        'logWordWps' = 'WPS Writer détecté, documents Word pris en charge'
        'logWordNo' = '[!!] Word/WPS introuvable, doc/docx ne peuvent pas être imprimés'
        'logExcelMs' = 'Microsoft Excel détecté, feuilles de calcul prises en charge'
        'logExcelWps' = 'WPS Tableurs détecté, feuilles de calcul prises en charge'
        'logExcelNo' = '[!!] Excel/WPS Tableurs introuvable, xls/xlsx ne peuvent pas être imprimés'
        'logPrinter' = 'Imprimante : {0}'
        'logSettings' = 'Paramètres : Copies={0}  Recto-verso={1}  Orientation={2}'
        'logStart' = 'Impression par lot démarrée, {0} fichier(s)'
        'logAbort' = 'Impression annulée par l''utilisateur'
        'logDuplexApplied' = 'Mode recto-verso appliqué : {0}'
        'logDuplexFail' = '[!!] Impossible de définir le mode recto-verso ({0}), impression avec les paramètres actuels'
        'logDuplexRestore' = 'Paramètre recto-verso de l''imprimante restauré : {0}'
        'logDuplexRestoreFail' = '[!!] Échec de la restauration du paramètre recto-verso : {0}'
        'logImage' = '[OK] Image {0}'
        'logText' = '[OK] Texte {0}'
        'logPdf' = '[OK] PDF {0}'
        'logWord' = '[OK] Word {0}'
        'logExcel' = '[OK] Excel {0}'
        'logFail' = '[!!] Échec de l''impression : {0}'
        'logSkip' = '[--] Format non pris en charge ignoré : {0} ({1})'
        'logDone' = 'Terminé : {0} réussi(s), {1} échoué(s)'
        'logPdfNeed' = 'PDF nécessite SumatraPDF : {0}'
        'logPdfTip' = 'Installez SumatraPDF (gratuit) et réessayez, ou ouvrez le PDF dans un navigateur et imprimez manuellement.'
        'errNoExcel' = 'Excel ou WPS Tableurs n''est pas installé, impossible d''imprimer les fichiers {0}'
        'errNoWord' = 'Word ou WPS Writer n''est pas installé, impossible d''imprimer les fichiers {0}'
        'errSwitchPrinter' = 'Impossible de passer à l''imprimante [{0}], vérifiez le nom de l''imprimante'
        'fileMissing' = 'manquant / inaccessible'
    }
    'es-ES' = @{
        'formTitle' = 'Herramienta de impresión por lotes - Versión local'
        'lblPrinter' = 'Impresora: '
        'lblLang' = 'Idioma: '
        'grpParam' = 'Configuración de impresión'
        'lblCopies' = 'Copias: '
        'lblDuplex' = 'Dúplex: '
        'lblOrient' = 'Orientación: '
        'lblTip' = 'Consejo: si su impresora no aparece, agréguela primero en la configuración de Windows y reinicie esta herramienta.'
        'btnAdd' = 'Agregar archivos'
        'btnRemove' = 'Quitar selección'
        'btnClear' = 'Vaciar lista'
        'lblCount' = '{0} archivo(s)'
        'btnPrint' = 'Iniciar impresión'
        'btnPrinting' = 'Imprimiendo...'
        'dlgTitle' = 'Seleccionar archivos para imprimir por lotes'
        'dlgFilter' = 'Archivos compatibles|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.txt;*.pdf;*.doc;*.docx;*.xls;*.xlsx|Imágenes|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|Texto|*.txt|PDF|*.pdf|Word|*.doc;*.docx|Excel|*.xls;*.xlsx|Todos los archivos|*.*'
        'dupOne' = 'Una cara'
        'dupLong' = 'Dúplex - borde largo'
        'dupShort' = 'Dúplex - borde corto'
        'orientAuto' = 'Automático'
        'orientPortrait' = 'Vertical'
        'orientLandscape' = 'Horizontal'
        'msgNoFiles' = 'Agregue primero los archivos para imprimir'
        'msgNoPrinter' = 'No se detectó ninguna impresora. Agregue una impresora en la configuración de Windows primero.'
        'msgTitle' = 'Aviso'
        'logTitle' = 'Herramienta de impresión por lotes - Versión local'
        'logPrinterCount' = 'Impresoras detectadas: {0}'
        'logNoPrinter' = '[!!] No se detectó ninguna impresora. Agregue una impresora en la configuración de Windows primero.'
        'logSumatraYes' = 'SumatraPDF detectado, impresión PDF silenciosa compatible'
        'logSumatraNo' = 'SumatraPDF no encontrado: los PDF se imprimirán manualmente (compatible tras instalarlo)'
        'logWordMs' = 'Microsoft Word detectado, documentos Word compatibles'
        'logWordWps' = 'WPS Writer detectado, documentos Word compatibles'
        'logWordNo' = '[!!] Word/WPS no encontrado, doc/docx no se pueden imprimir'
        'logExcelMs' = 'Microsoft Excel detectado, hojas de cálculo compatibles'
        'logExcelWps' = 'WPS Hojas de cálculo detectado, hojas de cálculo compatibles'
        'logExcelNo' = '[!!] Excel/WPS Hojas de cálculo no encontrado, xls/xlsx no se pueden imprimir'
        'logPrinter' = 'Impresora: {0}'
        'logSettings' = 'Configuración: Copias={0}  Dúplex={1}  Orientación={2}'
        'logStart' = 'Iniciando impresión por lotes, {0} archivo(s)'
        'logAbort' = 'Impresión cancelada por el usuario'
        'logDuplexApplied' = 'Modo dúplex aplicado: {0}'
        'logDuplexFail' = '[!!] No se pudo establecer el modo dúplex ({0}), imprimiendo con la configuración actual'
        'logDuplexRestore' = 'Configuración dúplex original restaurada: {0}'
        'logDuplexRestoreFail' = '[!!] No se pudo restaurar la configuración dúplex: {0}'
        'logImage' = '[OK] Imagen {0}'
        'logText' = '[OK] Texto {0}'
        'logPdf' = '[OK] PDF {0}'
        'logWord' = '[OK] Word {0}'
        'logExcel' = '[OK] Excel {0}'
        'logFail' = '[!!] Error de impresión: {0}'
        'logSkip' = '[--] Formato no compatible omitido: {0} ({1})'
        'logDone' = 'Listo: {0} con éxito, {1} con error'
        'logPdfNeed' = 'PDF requiere SumatraPDF: {0}'
        'logPdfTip' = 'Instale SumatraPDF (gratis) y vuelva a intentarlo, o abra el PDF en un navegador e imprima manualmente.'
        'errNoExcel' = 'Excel o WPS Hojas de cálculo no está instalado, no se pueden imprimir archivos {0}'
        'errNoWord' = 'Word o WPS Writer no está instalado, no se pueden imprimir archivos {0}'
        'errSwitchPrinter' = 'No se puede cambiar a la impresora [{0}], verifique el nombre de la impresora'
        'fileMissing' = 'faltante / inaccesible'
    }
}

# Current language: user override file first, then system UI language
# (Windows PowerShell 5.1's CurrentUICulture is unreliable, so use InstalledUICulture)
$script:LANG_LIST = @('zh-CN','en-US','zh-TW','ja-JP','ko-KR','de-DE','fr-FR','es-ES')
$script:LANG = 'en-US'
$langFile = Join-Path $PSScriptRoot 'lang.ini'
if (Test-Path $langFile) {
    $saved = (Get-Content $langFile -Raw -ErrorAction SilentlyContinue).Trim()
    if ($saved -in $script:LANG_LIST) { $script:LANG = $saved }
} else {
    $uiName = [System.Globalization.CultureInfo]::InstalledUICulture.Name
    if ($uiName -like 'zh-*') { $script:LANG = if ($uiName -in @('zh-TW','zh-HK','zh-MO')) { 'zh-TW' } else { 'zh-CN' } }
    elseif ($uiName -like 'ja-*') { $script:LANG = 'ja-JP' }
    elseif ($uiName -like 'ko-*') { $script:LANG = 'ko-KR' }
    elseif ($uiName -like 'de-*') { $script:LANG = 'de-DE' }
    elseif ($uiName -like 'fr-*') { $script:LANG = 'fr-FR' }
    elseif ($uiName -like 'es-*') { $script:LANG = 'es-ES' }
}


function T([string]$key) {
    $dict = $script:S[$script:LANG]
    if ($dict -and $dict.ContainsKey($key)) { return $dict[$key] }
    return $script:S['zh-CN'][$key]
}
function T2([string]$key, [object[]]$fmtArgs) {
    return [string]::Format((T $key), $fmtArgs)
}

# ---------------- Form ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Batch Print Tool'
$form.Size = New-Object System.Drawing.Size(760, 640)
$form.MinimumSize = New-Object System.Drawing.Size(680, 560)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)

# ---------------- Printer selection ----------------
$lblPrinter = New-Object System.Windows.Forms.Label
$lblPrinter.Location = New-Object System.Drawing.Point(14, 18)
$lblPrinter.AutoSize = $true

$cmbPrinter = New-Object System.Windows.Forms.ComboBox
$cmbPrinter.Location = New-Object System.Drawing.Point(80, 14)
$cmbPrinter.Size = New-Object System.Drawing.Size(360, 26)
$cmbPrinter.DropDownStyle = 'DropDownList'

# ---------------- Language selection ----------------
$lblLang = New-Object System.Windows.Forms.Label
$lblLang.Location = New-Object System.Drawing.Point(450, 18)
$lblLang.AutoSize = $true

$cmbLang = New-Object System.Windows.Forms.ComboBox
$cmbLang.Location = New-Object System.Drawing.Point(525, 14)
$cmbLang.Size = New-Object System.Drawing.Size(150, 26)
$cmbLang.DropDownStyle = 'DropDownList'
$cmbLang.Items.AddRange(@('简体中文','English','繁體中文','日本語','한국어','Deutsch','Français','Español'))
$cmbLang.add_SelectedIndexChanged({
    param($sender, $e)
    if ($cmbLang.SelectedIndex -ge 0 -and $cmbLang.SelectedIndex -lt $script:LANG_LIST.Count) {
        $script:LANG = $script:LANG_LIST[$cmbLang.SelectedIndex]
        try { Set-Content -Path (Join-Path $PSScriptRoot 'lang.ini') -Value $script:LANG -Encoding ASCII } catch {}
        Set-UIStrings
    }
})

# Enumerate printers
$printers = [System.Drawing.Printing.PrinterSettings]::InstalledPrinters
foreach ($p in $printers) { [void]$cmbPrinter.Items.Add([string]$p) }
if ($cmbPrinter.Items.Count -gt 0) { $cmbPrinter.SelectedIndex = 0 }

$form.Controls.Add($lblPrinter)
$form.Controls.Add($cmbPrinter)
$form.Controls.Add($lblLang)
$form.Controls.Add($cmbLang)

# ---------------- Print parameters ----------------
$grpParam = New-Object System.Windows.Forms.GroupBox
$grpParam.Location = New-Object System.Drawing.Point(14, 52)
$grpParam.Size = New-Object System.Drawing.Size(716, 88)

$lblCopies = New-Object System.Windows.Forms.Label
$lblCopies.Location = New-Object System.Drawing.Point(16, 28)
$lblCopies.AutoSize = $true

$numCopies = New-Object System.Windows.Forms.NumericUpDown
$numCopies.Location = New-Object System.Drawing.Point(60, 24)
$numCopies.Size = New-Object System.Drawing.Size(60, 24)
$numCopies.Minimum = 1
$numCopies.Maximum = 99
$numCopies.Value = 1

$lblDuplex = New-Object System.Windows.Forms.Label
$lblDuplex.Location = New-Object System.Drawing.Point(150, 28)
$lblDuplex.AutoSize = $true

$cmbDuplex = New-Object System.Windows.Forms.ComboBox
$cmbDuplex.Location = New-Object System.Drawing.Point(195, 24)
$cmbDuplex.Size = New-Object System.Drawing.Size(170, 26)
$cmbDuplex.DropDownStyle = 'DropDownList'

$lblOrient = New-Object System.Windows.Forms.Label
$lblOrient.Location = New-Object System.Drawing.Point(380, 28)
$lblOrient.AutoSize = $true

$cmbOrient = New-Object System.Windows.Forms.ComboBox
$cmbOrient.Location = New-Object System.Drawing.Point(440, 24)
$cmbOrient.Size = New-Object System.Drawing.Size(120, 26)
$cmbOrient.DropDownStyle = 'DropDownList'

$lblTip = New-Object System.Windows.Forms.Label
$lblTip.Location = New-Object System.Drawing.Point(16, 60)
$lblTip.Size = New-Object System.Drawing.Size(680, 20)
$lblTip.ForeColor = [System.Drawing.Color]::Gray
$lblTip.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 8)

$grpParam.Controls.Add($lblCopies);  $grpParam.Controls.Add($numCopies)
$grpParam.Controls.Add($lblDuplex);  $grpParam.Controls.Add($cmbDuplex)
$grpParam.Controls.Add($lblOrient);  $grpParam.Controls.Add($cmbOrient)
$grpParam.Controls.Add($lblTip)
$form.Controls.Add($grpParam)

# ---------------- File list ----------------
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Location = New-Object System.Drawing.Point(14, 152)
$btnAdd.Size = New-Object System.Drawing.Size(96, 30)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Location = New-Object System.Drawing.Point(118, 152)
$btnRemove.Size = New-Object System.Drawing.Size(96, 30)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(222, 152)
$btnClear.Size = New-Object System.Drawing.Size(96, 30)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Location = New-Object System.Drawing.Point(330, 158)
$lblCount.AutoSize = $true
$lblCount.ForeColor = [System.Drawing.Color]::Gray

$listFiles = New-Object System.Windows.Forms.ListBox
$listFiles.Location = New-Object System.Drawing.Point(14, 190)
$listFiles.Size = New-Object System.Drawing.Size(716, 220)
$listFiles.SelectionMode = 'MultiExtended'
$listFiles.HorizontalScrollbar = $true

$form.Controls.Add($btnAdd)
$form.Controls.Add($btnRemove)
$form.Controls.Add($btnClear)
$form.Controls.Add($lblCount)
$form.Controls.Add($listFiles)

# ---------------- Log area ----------------
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(14, 430)
$txtLog.Size = New-Object System.Drawing.Size(716, 110)
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)

$form.Controls.Add($txtLog)

# ---------------- Print button ----------------
$btnPrint = New-Object System.Windows.Forms.Button
$btnPrint.Location = New-Object System.Drawing.Point(620, 552)
$btnPrint.Size = New-Object System.Drawing.Size(110, 40)
$btnPrint.BackColor = [System.Drawing.Color]::FromArgb(61, 90, 241)
$btnPrint.ForeColor = [System.Drawing.Color]::White
$btnPrint.FlatStyle = 'Flat'
$btnPrint.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)

$form.Controls.Add($btnPrint)

# ---------------- Functions ----------------
function Add-Log([string]$msg) {
    $txtLog.AppendText($msg + "`r`n")
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Apply all UI texts for the current language (also rebuilds combo items)
function Set-UIStrings {
    $form.Text = T 'formTitle'
    $lblPrinter.Text = T 'lblPrinter'
    $lblLang.Text = T 'lblLang'
    $grpParam.Text = T 'grpParam'
    $lblCopies.Text = T 'lblCopies'
    $lblDuplex.Text = T 'lblDuplex'
    $lblOrient.Text = T 'lblOrient'
    $lblTip.Text = T 'lblTip'
    $btnAdd.Text = T 'btnAdd'
    $btnRemove.Text = T 'btnRemove'
    $btnClear.Text = T 'btnClear'
    $btnPrint.Text = T 'btnPrint'

    $dupSel = $cmbDuplex.SelectedIndex
    $cmbDuplex.Items.Clear()
    $cmbDuplex.Items.AddRange(@((T 'dupOne'), (T 'dupLong'), (T 'dupShort')))
    if ($dupSel -ge 0 -and $dupSel -lt $cmbDuplex.Items.Count) { $cmbDuplex.SelectedIndex = $dupSel }
    else { $cmbDuplex.SelectedIndex = 0 }

    $oriSel = $cmbOrient.SelectedIndex
    $cmbOrient.Items.Clear()
    $cmbOrient.Items.AddRange(@((T 'orientAuto'), (T 'orientPortrait'), (T 'orientLandscape')))
    if ($oriSel -ge 0 -and $oriSel -lt $cmbOrient.Items.Count) { $cmbOrient.SelectedIndex = $oriSel }
    else { $cmbOrient.SelectedIndex = 0 }

    Refresh-List
}

function Add-Files {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Multiselect = $true
    $dlg.Title = T 'dlgTitle'
    $dlg.Filter = T 'dlgFilter'
    if ($dlg.ShowDialog() -eq 'OK') {
        foreach ($f in $dlg.FileNames) {
            if (-not $script:fileList.Contains($f)) { [void]$script:fileList.Add($f) }
        }
        Refresh-List
    }
}

function Refresh-List {
    $listFiles.Items.Clear()
    foreach ($f in $script:fileList) {
        $label = [System.IO.Path]::GetFileName($f)
        try {
            $item = Get-Item $f -ErrorAction Stop
            $label += '   [' + ($item.Length / 1KB).ToString('0.0') + ' KB]'
        } catch {
            $label += '   [' + (T 'fileMissing') + ']'
        }
        [void]$listFiles.Items.Add($label)
    }
    $lblCount.Text = T2 'lblCount' @($script:fileList.Count)
}

# Find SumatraPDF (for PDF printing)
function Find-SumatraPdf {
    $candidates = @(
        "$env:ProgramFiles\SumatraPDF\SumatraPDF.exe",
        "${env:ProgramFiles(x86)}\SumatraPDF\SumatraPDF.exe",
        "$env:LOCALAPPDATA\SumatraPDF\SumatraPDF.exe",
        "$env:ProgramData\SumatraPDF\SumatraPDF.exe"
    )
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    return $null
}

# Print one image (centered, scaled to page; copies/duplex/orientation)
function Print-ImageFile([string]$path, [string]$printer, [int]$copies, [int]$dupIdx, [bool]$landscape, [bool]$orientAuto) {
    $img = $null
    try {
        $img = [System.Drawing.Image]::FromFile($path)
        $pd = New-Object System.Drawing.Printing.PrintDocument
        $pd.PrinterSettings.PrinterName = $printer
        $pd.PrinterSettings.Copies = $copies
        if ($dupIdx -eq 1) { $pd.PrinterSettings.Duplex = [System.Drawing.Printing.Duplex]::Horizontal }
        elseif ($dupIdx -eq 2) { $pd.PrinterSettings.Duplex = [System.Drawing.Printing.Duplex]::Vertical }
        else { $pd.PrinterSettings.Duplex = [System.Drawing.Printing.Duplex]::Simplex }
        if ($orientAuto) { $landscape = ($img.Width -gt $img.Height) }
        $pd.DefaultPageSettings.Landscape = $landscape

        $script:curImg = $img
        $pd.add_PrintPage({
            param($sender, $evt)
            $img = $script:curImg
            $page = $evt.PageBounds
            $scale = [Math]::Min($page.Width / $img.Width, $page.Height / $img.Height)
            $w = [int]($img.Width * $scale)
            $h = [int]($img.Height * $scale)
            $x = [int](($page.Width - $w) / 2)
            $y = [int](($page.Height - $h) / 2)
            $evt.Graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $evt.Graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $evt.Graphics.DrawImage($img, $x, $y, $w, $h)
            $evt.HasMorePages = $false
        })
        $pd.Print()
        Add-Log (T2 'logImage' @([System.IO.Path]::GetFileName($path)))
    } finally {
        if ($img) { $img.Dispose() }
    }
}

# Print one text file (paginated line output)
function Print-TextFile([string]$path, [string]$printer, [int]$copies, [int]$dupIdx, [bool]$landscape, [bool]$orientAuto) {
    $content = ''
    try {
        $bytes = [System.IO.File]::ReadAllBytes($path)
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            $content = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
        } else {
            $content = [System.Text.Encoding]::UTF8.GetString($bytes)
            if ($content.Contains([char]0xFFFD)) {
                $content = [System.Text.Encoding]::GetEncoding(936).GetString($bytes)
            }
        }
    } catch {
        $content = [System.IO.File]::ReadAllText($path)
    }
    $lines = $content -split "`r`n|`n"

    $pd = New-Object System.Drawing.Printing.PrintDocument
    $pd.PrinterSettings.PrinterName = $printer
    $pd.PrinterSettings.Copies = $copies
    if ($dupIdx -eq 1) { $pd.PrinterSettings.Duplex = [System.Drawing.Printing.Duplex]::Horizontal }
    elseif ($dupIdx -eq 2) { $pd.PrinterSettings.Duplex = [System.Drawing.Printing.Duplex]::Vertical }
    else { $pd.PrinterSettings.Duplex = [System.Drawing.Printing.Duplex]::Simplex }
    if ($orientAuto) { $landscape = $false }
    $pd.DefaultPageSettings.Landscape = $landscape

    $script:txtLines = $lines
    $script:lineIdx = 0
    $pd.add_PrintPage({
        param($sender, $evt)
        $font = New-Object System.Drawing.Font('Microsoft YaHei', 12)
        $lineHeight = [int]($font.GetHeight($evt.Graphics) + 4)
        $margin = 40
        $x = $margin
        $y = $margin
        $maxY = $evt.PageBounds.Height - $margin
        $drawn = 0
        while ($script:lineIdx -lt $script:txtLines.Count -and $y + $lineHeight -le $maxY) {
            $evt.Graphics.DrawString($script:txtLines[$script:lineIdx], $font, [System.Drawing.Brushes]::Black, $x, $y)
            $script:lineIdx++
            $y += $lineHeight
            $drawn++
        }
        $evt.HasMorePages = ($script:lineIdx -lt $script:txtLines.Count)
        if ($drawn -eq 0) { $evt.HasMorePages = $false }
    })
    $pd.Print()
    Add-Log (T2 'logText' @([System.IO.Path]::GetFileName($path)))
}

# Print PDF via SumatraPDF command line
function Print-PdfFile([string]$path, [string]$printer, [int]$copies, [int]$dupIdx, [bool]$landscape, [bool]$orientAuto) {
    $sumatra = Find-SumatraPdf
    if (-not $sumatra) {
        Add-Log (T2 'logPdfNeed' @($path))
        Add-Log (T 'logPdfTip')
        return
    }
    $settings = 'duplex=off'
    if ($dupIdx -eq 1) { $settings = 'duplex=long' }
    elseif ($dupIdx -eq 2) { $settings = 'duplex=short' }
    if ($landscape -and -not $orientAuto) { $settings += ',landscape' }
    if ($copies -gt 1) { $settings += ",copies=$copies" }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $sumatra
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.Arguments = "-print-to `"$printer`" -print-settings `"$settings`" `"$path`""
    $proc = [System.Diagnostics.Process]::Start($psi)
    if (-not $proc.WaitForExit(15000)) {
        try { $proc.Kill() } catch {}
        throw 'SumatraPDF 打印超时'
    }
    if ($proc.ExitCode -ne 0) { throw "SumatraPDF 退出码 $($proc.ExitCode)" }
    Add-Log (T2 'logPdf' @([System.IO.Path]::GetFileName($path)))
}

# Print one Word/Excel document (Microsoft Office first, fallback to WPS)
# Duplex: handled by the main flow via printer default mode switch
# Printer selection: switch the system default printer (SetDefaultPrinter),
# which is more reliable than Office's ActivePrinter on all systems/languages
function Print-OfficeDoc([string]$path, [string]$printer, [int]$copies, [bool]$landscape, [bool]$orientAuto) {
    $ext = [System.IO.Path]::GetExtension($path).ToLower()
    $isExcel = ($ext -in '.xls', '.xlsx')

    $app = $null
    $doc = $null
    $oldDefault = Get-SystemDefaultPrinter
    try {
        if ($isExcel) {
            try { $app = New-Object -ComObject Excel.Application -ErrorAction Stop }
            catch { try { $app = New-Object -ComObject KET.Application -ErrorAction Stop } catch { try { $app = New-Object -ComObject Et.Application -ErrorAction Stop } catch { throw (T2 'errNoExcel' @($ext)) } } }
        } else {
            try { $app = New-Object -ComObject Word.Application -ErrorAction Stop }
            catch { try { $app = New-Object -ComObject KWps.Application -ErrorAction Stop } catch { throw (T2 'errNoWord' @($ext)) } }
        }

        $app.Visible = $false
        if ($isExcel) { $app.DisplayAlerts = $false } else { $app.DisplayAlerts = 0 }

        if ($isExcel) { $doc = $app.Workbooks.Open($path, $false, $true) }
        else { $doc = $app.Documents.Open($path, $false, $true) }

        # Orientation: Auto keeps the document's own orientation (prevents splitting wide tables)
        if (-not $orientAuto) {
            if ($isExcel) {
                # xlPortrait=1, xlLandscape=2
                foreach ($ws in $doc.Worksheets) {
                    if ($landscape) { $ws.PageSetup.Orientation = 2 } else { $ws.PageSetup.Orientation = 1 }
                }
            } else {
                # wdOrientPortrait=0, wdOrientLandscape=1
                if ($landscape) { $doc.PageSetup.Orientation = 1 } else { $doc.PageSetup.Orientation = 0 }
            }
        }

        # Switch to the chosen printer, then print using the default printer
        if ($oldDefault -ne $printer) {
            $ok = [PrintApi]::SetDefaultPrinter($printer)
            if (-not $ok) { throw (T2 'errSwitchPrinter' @($printer)) }
            if ((Get-SystemDefaultPrinter) -ne $printer) { throw (T2 'errSwitchPrinter' @($printer)) }
        }

        $missing = [System.Reflection.Missing]::Value
        if ($isExcel) {
            # (From, To, Copies, Collate, ActivePrinter, PrintToFile, Preview, PrintToFileName, IgnorePrintAreas)
            $doc.PrintOut($missing, $missing, $copies, $false, $missing, $false, $false, $missing, $true)
        } else {
            # Word positional args; ActivePrinter omitted -> use system default
            $doc.PrintOut($false, $false, 0, $missing, $missing, $missing, $missing, $copies, $missing, $missing, $false, $true, $missing, $false)
        }

        $doc.Close($false)
        $app.Quit()
        $app = $null
        $kind = if ($isExcel) { T2 'logExcel' @([System.IO.Path]::GetFileName($path)) } else { T2 'logWord' @([System.IO.Path]::GetFileName($path)) }
        Add-Log $kind
    } catch {
        if ($doc) { try { $doc.Close($false) } catch {} }
        if ($app) { try { $app.Quit() } catch {} }
        throw $_
    } finally {
        if ($oldDefault) { [PrintApi]::SetDefaultPrinter($oldDefault) | Out-Null }
    }
}

# Main print flow
function Start-Printing {
    if ($script:fileList.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show((T 'msgNoFiles'), (T 'msgTitle'))
        return
    }
    if ($cmbPrinter.SelectedItem -eq $null) {
        [System.Windows.Forms.MessageBox]::Show((T 'msgNoPrinter'), (T 'msgTitle'))
        return
    }

    $printer = [string]$cmbPrinter.SelectedItem
    $copies  = [int]$numCopies.Value
    $dupIdx  = $cmbDuplex.SelectedIndex
    $oriIdx  = $cmbOrient.SelectedIndex
    $landscape = ($oriIdx -eq 2)
    $orientAuto = ($oriIdx -eq 0)

    switch ($dupIdx) {
        1 { $targetMode = 'TwoSidedLongEdge';  $dupLabel = T 'dupLong' }
        2 { $targetMode = 'TwoSidedShortEdge'; $dupLabel = T 'dupShort' }
        default { $targetMode = 'OneSided';    $dupLabel = T 'dupOne' }
    }
    $oriLabel = switch ($oriIdx) { 0 { T 'orientAuto' } 1 { T 'orientPortrait' } 2 { T 'orientLandscape' } }

    $btnPrint.Enabled = $false
    $btnPrint.Text = T 'btnPrinting'
    Add-Log '=============================================='
    Add-Log (T2 'logPrinter' @($printer))
    Add-Log (T2 'logSettings' @($copies, $dupLabel, $oriLabel))
    Add-Log (T2 'logStart' @($script:fileList.Count))
    Add-Log '----------------------------------------------'

    # Before printing: temporarily switch the printer to the chosen duplex mode
    # (Word/Excel print with the printer's default settings)
    $savedDuplexMode = $null
    $duplexChanged = $false
    try {
        $cfg = Get-PrintConfiguration -PrinterName $printer -ErrorAction Stop
        $savedDuplexMode = [string]$cfg.DuplexingMode
        if ($savedDuplexMode -ne $targetMode) {
            Set-PrintConfiguration -PrinterName $printer -DuplexingMode $targetMode
            $duplexChanged = $true
        }
        Add-Log (T2 'logDuplexApplied' @($dupLabel))
    } catch {
        Add-Log (T2 'logDuplexFail' @($_.Exception.Message))
    }

    $okCount = 0; $failCount = 0
    foreach ($path in $script:fileList) {
        if ($script:abort) { Add-Log (T 'logAbort'); break }
        $ext = [System.IO.Path]::GetExtension($path).ToLower()
        try {
            switch ($ext) {
                { $_ -in '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff' } { Print-ImageFile $path $printer $copies $dupIdx $landscape $orientAuto; $okCount++ }
                '.txt' { Print-TextFile $path $printer $copies $dupIdx $landscape $orientAuto; $okCount++ }
                '.pdf' { Print-PdfFile $path $printer $copies $dupIdx $landscape $orientAuto; $okCount++ }
                { $_ -in '.doc', '.docx', '.xls', '.xlsx' } { Print-OfficeDoc $path $printer $copies $landscape $orientAuto; $okCount++ }
                default { Add-Log (T2 'logSkip' @([System.IO.Path]::GetFileName($path), $ext)) }
            }
        } catch {
            Add-Log (T2 'logFail' @([System.IO.Path]::GetFileName($path)))
            Add-Log "     $($_.Exception.Message)"
            $failCount++
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Restore the printer's original duplex setting
    if ($duplexChanged) {
        try { Set-PrintConfiguration -PrinterName $printer -DuplexingMode $savedDuplexMode; Add-Log (T2 'logDuplexRestore' @($savedDuplexMode)) }
        catch { Add-Log (T2 'logDuplexRestoreFail' @($_.Exception.Message)) }
    }

    Add-Log '----------------------------------------------'
    Add-Log (T2 'logDone' @($okCount, $failCount))
    Add-Log '=============================================='
    $btnPrint.Enabled = $true
    $btnPrint.Text = T 'btnPrint'
}

# ---------------- Event bindings ----------------
$btnAdd.add_Click({ Add-Files })
$btnRemove.add_Click({
    $selected = $listFiles.SelectedIndices
    for ($i = $selected.Count - 1; $i -ge 0; $i--) {
        $script:fileList.RemoveAt($selected[$i])
    }
    Refresh-List
})
$btnClear.add_Click({
    $script:fileList.Clear()
    Refresh-List
    $txtLog.Clear()
})
$btnPrint.add_Click({ Start-Printing })


# Drag & drop files into the list
$listFiles.AllowDrop = $true
$listFiles.add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = 'Copy' }
})
$listFiles.add_DragDrop({
    param($sender, $e)
    $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    foreach ($f in $files) {
        if (Test-Path $f -PathType Leaf -and -not $script:fileList.Contains($f)) { [void]$script:fileList.Add($f) }
    }
    Refresh-List
})

# Abort printing on form close
$form.add_FormClosing({
    param($sender, $e)
    $script:abort = $true
})

# ---------------- Startup ----------------
# Apply language: file / system detection first, user can override in UI
$cmbLang.SelectedIndex = [Math]::Max(0, [Array]::IndexOf($script:LANG_LIST, $script:LANG))
Set-UIStrings

Add-Log (T 'logTitle')
Add-Log (T2 'logPrinterCount' @($printers.Count))
if ($printers.Count -eq 0) { Add-Log (T 'logNoPrinter') }
$sum = Find-SumatraPdf
if ($sum) { Add-Log (T 'logSumatraYes') }
else { Add-Log (T 'logSumatraNo') }
try { $w = New-Object -ComObject Word.Application -ErrorAction Stop; Add-Log (T 'logWordMs'); $w.Quit() }
catch { try { $w2 = New-Object -ComObject KWps.Application -ErrorAction Stop; Add-Log (T 'logWordWps'); $w2.Quit() } catch { Add-Log (T 'logWordNo') } }
try { $e = New-Object -ComObject Excel.Application -ErrorAction Stop; Add-Log (T 'logExcelMs'); $e.Quit() }
catch { try { $e2 = New-Object -ComObject KET.Application -ErrorAction Stop; Add-Log (T 'logExcelWps'); $e2.Quit() } catch { Add-Log (T 'logExcelNo') } }
Add-Log '----------------------------------------------'

$form.ShowDialog()
