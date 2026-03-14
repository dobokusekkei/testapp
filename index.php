<?php
// ==========================================
// 1. システム設定 ＆ データベース初期化
// ==========================================
ini_set('display_errors', 1);
error_reporting(E_ALL);
ini_set('memory_limit', '1024M'); 

$autoload_path = realpath(__DIR__ . '/../vendor/autoload.php');
if (!$autoload_path || !file_exists($autoload_path)) {
    die("<h3>システムエラー: vendor/autoload.php が見つかりません。</h3>");
}
require $autoload_path;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

$default_safety = "・健康状態の悪い者は、作業に従事しないこと。作業中に具合が悪くなった場合には、作業指揮者に必ず連絡すること。\n・作業時は、保護帽・安全靴を必ず着用して作業を行うこと。\n・作業中及び歩行中の禁煙を徹底すること。喫煙は、決められた場所にて行うこと。\n・通勤等で使用する車両の運行では、交通法規を遵守し交通災害の防止に努めること。また、使用する車両は定期検査等の法令に則った検査を実施していること。\n・点検工具の置忘れに注意すること。\n・作業員は、指示された作業内容、モーターカー運行状況、待避を厳守すること。\n・作業員は、線路閉鎖が完了し、作業指揮者からの指示があるまで線路内への立ち入りを禁ずる。\n\n\n\n\n\n\n\n\n";

// ★ 祝日リスト（必要に応じて自由に追加・変更してください）
$holidays = [
    '2026-01-01', '2026-01-12', '2026-02-11', '2026-02-23', '2026-03-20',
    '2026-04-29', '2026-05-03', '2026-05-04', '2026-05-05', '2026-05-06',
    '2026-07-20', '2026-08-11', '2026-09-21', '2026-09-22', '2026-09-23',
    '2026-10-12', '2026-11-03', '2026-11-23'
];

function isHolidayOrWeekend($dateString, $holidays) {
    $timestamp = strtotime($dateString);
    $w = date('w', $timestamp); // 0:日, 6:土
    if ($w == 0 || $w == 6) return true;
    if (in_array(date('Y-m-d', $timestamp), $holidays)) return true;
    return false;
}

$dbFile = __DIR__ . '/anzen_db.sqlite';
$pdo = new PDO('sqlite:' . $dbFile);
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

$pdo->exec("CREATE TABLE IF NOT EXISTS personnel (id INTEGER PRIMARY KEY AUTOINCREMENT, type TEXT, company TEXT, name TEXT, phone TEXT)");
$pdo->exec("CREATE TABLE IF NOT EXISTS saved_plans (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, form_data TEXT, created_at DATETIME DEFAULT CURRENT_TIMESTAMP)");
$pdo->exec("CREATE TABLE IF NOT EXISTS safety_templates (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, content TEXT)");
$pdo->exec("CREATE TABLE IF NOT EXISTS team_settings (id INTEGER PRIMARY KEY AUTOINCREMENT, team_name TEXT, contact1_name TEXT, contact1_phone TEXT, contact2_name TEXT, contact2_phone TEXT)");

// DBアップデート
$rs = $pdo->query("PRAGMA table_info(team_settings)");
$columns = [];
foreach($rs as $r) { $columns[] = $r['name']; }
if (!in_array('group_name', $columns)) {
    $pdo->exec("ALTER TABLE team_settings ADD COLUMN group_name TEXT DEFAULT ''");
    $pdo->exec("ALTER TABLE team_settings ADD COLUMN group_leader_name TEXT DEFAULT ''");
    $pdo->exec("ALTER TABLE team_settings ADD COLUMN group_leader_phone TEXT DEFAULT ''");
}

$stmt = $pdo->query("SELECT COUNT(*) FROM safety_templates");
if ($stmt->fetchColumn() == 0) {
    $ins = $pdo->prepare("INSERT INTO safety_templates (title, content) VALUES (?, ?)");
    $ins->execute(["軌道内夜間作業", trim($default_safety)]);
}
$stmt = $pdo->query("SELECT COUNT(*) FROM team_settings");
if ($stmt->fetchColumn() == 0) {
    $teams = ['用地Ｔ', '設計1Ｔ', '設計2Ｔ', '設計3Ｔ', '設計4Ｔ'];
    $ins = $pdo->prepare("INSERT INTO team_settings (group_name, group_leader_name, group_leader_phone, team_name, contact1_name, contact1_phone, contact2_name, contact2_phone) VALUES ('', '', '', ?, '', '', '', '')");
    foreach ($teams as $t) { $ins->execute([$t]); }
}

// ==========================================
// 2. AJAXリクエスト処理 (DB連携)
// ==========================================
if (isset($_POST['ajax_action'])) {
    header('Content-Type: application/json; charset=utf-8');
    try {
        if ($_POST['ajax_action'] === 'save_person') {
            $stmt = $pdo->prepare("INSERT INTO personnel (type, company, name, phone) VALUES (?, ?, ?, ?)");
            $stmt->execute([$_POST['type'], $_POST['company'] ?? '', $_POST['name'], $_POST['phone']]);
            echo json_encode(['status' => 'success']);
        }
        elseif ($_POST['ajax_action'] === 'delete_person') {
            $stmt = $pdo->prepare("DELETE FROM personnel WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            echo json_encode(['status' => 'success']);
        }
        elseif ($_POST['ajax_action'] === 'import_csv') {
            $data = json_decode($_POST['csv_data'], true);
            $stmt = $pdo->prepare("INSERT INTO personnel (type, company, name, phone) VALUES (?, ?, ?, ?)");
            $pdo->beginTransaction();
            foreach ($data as $row) {
                if (count($row) >= 4) {
                    $type = trim($row[0]) === 'our' ? 'our' : 'partner';
                    $company = trim($row[1]);
                    $name = trim($row[2]);
                    $phone = trim($row[3]);
                    if ($name !== '') $stmt->execute([$type, $company, $name, $phone]);
                }
            }
            $pdo->commit();
            echo json_encode(['status' => 'success']);
        }
        elseif ($_POST['ajax_action'] === 'get_personnel') {
            $stmt = $pdo->query("SELECT * FROM personnel ORDER BY company, name");
            echo json_encode($stmt->fetchAll(PDO::FETCH_ASSOC));
        }
        elseif ($_POST['ajax_action'] === 'save_plan') {
            $title = trim($_POST['title']);
            if (preg_match('/^(.*)_(\d+)$/', $title, $matches)) {
                $base_title = $matches[1];
                $counter = (int)$matches[2];
            } else {
                $base_title = $title;
                $counter = 2;
            }
            $check = $pdo->prepare("SELECT id FROM saved_plans WHERE title = ?");
            $check->execute([$title]);
            if ($check->fetch()) {
                while (true) {
                    $new_title = $base_title . '_' . $counter;
                    $check->execute([$new_title]);
                    if (!$check->fetch()) {
                        $title = $new_title;
                        break;
                    }
                    $counter++;
                }
            }
            $stmt = $pdo->prepare("INSERT INTO saved_plans (title, form_data) VALUES (?, ?)");
            $stmt->execute([$title, $_POST['form_data']]);
            echo json_encode(['status' => 'success', 'saved_title' => $title, 'id' => $pdo->lastInsertId()]);
        }
        elseif ($_POST['ajax_action'] === 'overwrite_plan') {
            $stmt = $pdo->prepare("UPDATE saved_plans SET title = ?, form_data = ? WHERE id = ?");
            $stmt->execute([$_POST['title'], $_POST['form_data'], $_POST['id']]);
            echo json_encode(['status' => 'success', 'saved_title' => $_POST['title'], 'id' => $_POST['id']]);
        }
        elseif ($_POST['ajax_action'] === 'update_plan_title') {
            $stmt = $pdo->prepare("UPDATE saved_plans SET title = ? WHERE id = ?");
            $stmt->execute([$_POST['title'], $_POST['id']]);
            echo json_encode(['status' => 'success']);
        }
        elseif ($_POST['ajax_action'] === 'get_plans') {
            $stmt = $pdo->query("SELECT id, title, created_at FROM saved_plans ORDER BY id DESC");
            echo json_encode($stmt->fetchAll(PDO::FETCH_ASSOC));
        }
        elseif ($_POST['ajax_action'] === 'get_plans_all') {
            $stmt = $pdo->query("SELECT id, title, created_at, form_data FROM saved_plans ORDER BY id DESC");
            echo json_encode($stmt->fetchAll(PDO::FETCH_ASSOC));
        }
        elseif ($_POST['ajax_action'] === 'load_plan') {
            $stmt = $pdo->prepare("SELECT form_data FROM saved_plans WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            echo json_encode($stmt->fetch(PDO::FETCH_ASSOC));
        }
        elseif ($_POST['ajax_action'] === 'delete_plan') {
            $stmt = $pdo->prepare("DELETE FROM saved_plans WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            echo json_encode(['status' => 'success']);
        }
        elseif ($_POST['ajax_action'] === 'get_templates') {
            $stmt = $pdo->query("SELECT * FROM safety_templates ORDER BY id");
            echo json_encode($stmt->fetchAll(PDO::FETCH_ASSOC));
        }
        elseif ($_POST['ajax_action'] === 'save_template') {
            if (!empty($_POST['id'])) {
                $stmt = $pdo->prepare("UPDATE safety_templates SET title=?, content=? WHERE id=?");
                $stmt->execute([$_POST['title'], $_POST['content'], $_POST['id']]);
            } else {
                $stmt = $pdo->prepare("INSERT INTO safety_templates (title, content) VALUES (?, ?)");
                $stmt->execute([$_POST['title'], $_POST['content']]);
            }
            echo json_encode(['status' => 'success']);
        }
        elseif ($_POST['ajax_action'] === 'delete_template') {
            $stmt = $pdo->prepare("DELETE FROM safety_templates WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            echo json_encode(['status' => 'success']);
        }
        elseif ($_POST['ajax_action'] === 'get_teams') {
            $stmt = $pdo->query("SELECT * FROM team_settings ORDER BY id");
            echo json_encode($stmt->fetchAll(PDO::FETCH_ASSOC));
        }
        elseif ($_POST['ajax_action'] === 'save_team') {
            if (!empty($_POST['id'])) {
                $stmt = $pdo->prepare("UPDATE team_settings SET group_name=?, group_leader_name=?, group_leader_phone=?, team_name=?, contact1_name=?, contact1_phone=?, contact2_name=?, contact2_phone=? WHERE id=?");
                $stmt->execute([
                    $_POST['group_name'], $_POST['group_leader_name'], $_POST['group_leader_phone'],
                    $_POST['team_name'], $_POST['contact1_name'], $_POST['contact1_phone'], $_POST['contact2_name'], $_POST['contact2_phone'], $_POST['id']
                ]);
            } else {
                $stmt = $pdo->prepare("INSERT INTO team_settings (group_name, group_leader_name, group_leader_phone, team_name, contact1_name, contact1_phone, contact2_name, contact2_phone) VALUES (?, ?, ?, ?, ?, ?, ?, ?)");
                $stmt->execute([
                    $_POST['group_name'], $_POST['group_leader_name'], $_POST['group_leader_phone'],
                    $_POST['team_name'], $_POST['contact1_name'], $_POST['contact1_phone'], $_POST['contact2_name'], $_POST['contact2_phone']
                ]);
            }
            echo json_encode(['status' => 'success']);
        }
        elseif ($_POST['ajax_action'] === 'delete_team') {
            $stmt = $pdo->prepare("DELETE FROM team_settings WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            echo json_encode(['status' => 'success']);
        }
    } catch (Exception $e) {
        if ($pdo->inTransaction()) $pdo->rollBack();
        echo json_encode(['status' => 'error', 'message' => $e->getMessage()]);
    }
    exit;
}

// ==========================================
// ★ 外業管理表のExcel出力処理
// ==========================================
if (isset($_POST['export_gaigyo_excel'])) {
    try {
        $data = json_decode($_POST['gaigyo_data'], true);
        
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('外業管理表');
        
        // ヘッダーセット
        $headers = ['No.', '作業日', '曜日', '業務名', '作業指揮者', '携帯番号', '作業員', '昼夜別', '時間帯', '夜達番号等', '関連夜達留変等', '場所', '業者名①', '作業責任者', '業者携帯', '人数', '列監', '整理員', '備考'];
        $sheet->fromArray($headers, null, 'A1');
        
        // ヘッダーの装飾
        $headerStyle = [
            'font' => ['bold' => true, 'color' => ['argb' => 'FFFFFFFF']],
            'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['argb' => 'FF005A9E']],
            'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
            'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]]
        ];
        $sheet->getStyle('A1:S1')->applyFromArray($headerStyle);
        
        // データセット (すべて文字列としてセット)
        $rowNum = 2;
        if (is_array($data)) {
            foreach($data as $row) {
                $colIdx = 1;
                foreach ($row as $cellValue) {
                    $colLetter = Coordinate::stringFromColumnIndex($colIdx);
                    $sheet->setCellValueExplicit($colLetter . $rowNum, (string)$cellValue, DataType::TYPE_STRING);
                    $colIdx++;
                }
                
                // 罫線と折り返し設定
                $sheet->getStyle('A'.$rowNum.':S'.$rowNum)->applyFromArray([
                    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]]
                ]);
                $sheet->getStyle('D'.$rowNum)->getAlignment()->setWrapText(true); // 業務名
                $sheet->getStyle('G'.$rowNum)->getAlignment()->setWrapText(true); // 作業員
                $sheet->getStyle('L'.$rowNum)->getAlignment()->setWrapText(true); // 場所
                $sheet->getStyle('M'.$rowNum)->getAlignment()->setWrapText(true); // 業者名①
                $sheet->getStyle('S'.$rowNum)->getAlignment()->setWrapText(true); // 備考
                $rowNum++;
            }
        }
        
        // ★ センタリング設定
        if ($rowNum > 2) {
            $centerCols = ['A', 'B', 'C', 'F', 'H', 'I', 'J', 'O', 'P', 'Q', 'R'];
            foreach($centerCols as $c) {
                $sheet->getStyle($c.'2:'.$c.($rowNum-1))->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            }

            // ★ 指定列のフォントサイズを9に縮小
            $smallFontCols = ['J', 'K', 'R']; // 夜達番号等, 関連夜達留変等, 整理員
            foreach($smallFontCols as $c) {
                $sheet->getStyle($c.'2:'.$c.($rowNum-1))->getFont()->setSize(9);
            }
        }

        // ★ 列幅の最適化
        $sheet->getColumnDimension('A')->setAutoSize(false)->setWidth(4.5);  // No.
        $sheet->getColumnDimension('B')->setAutoSize(false)->setWidth(11.5); // 作業日
        $sheet->getColumnDimension('C')->setAutoSize(false)->setWidth(5.5);  // 曜日
        
        $sheet->getColumnDimension('D')->setAutoSize(false)->setWidth(45);   // 業務名
        
        $sheet->getColumnDimension('E')->setAutoSize(true);                  // 作業指揮者
        $sheet->getColumnDimension('F')->setAutoSize(false)->setWidth(14.5); // 携帯番号
        $sheet->getColumnDimension('G')->setAutoSize(false)->setWidth(20);   // 作業員
        $sheet->getColumnDimension('H')->setAutoSize(false)->setWidth(8);    // 昼夜別
        $sheet->getColumnDimension('I')->setAutoSize(false)->setWidth(12.5); // 時間帯
        $sheet->getColumnDimension('J')->setAutoSize(false)->setWidth(11);   // 夜達番号等
        $sheet->getColumnDimension('K')->setAutoSize(false)->setWidth(15);   // 関連夜達留変等
        $sheet->getColumnDimension('L')->setAutoSize(false)->setWidth(20);   // 場所
        $sheet->getColumnDimension('M')->setAutoSize(false)->setWidth(18);   // 業者名①
        $sheet->getColumnDimension('N')->setAutoSize(true);                  // 作業責任者
        $sheet->getColumnDimension('O')->setAutoSize(false)->setWidth(14.5); // 業者携帯
        $sheet->getColumnDimension('P')->setAutoSize(false)->setWidth(5.5);  // 人数
        $sheet->getColumnDimension('Q')->setAutoSize(false)->setWidth(5.5);  // 列監
        $sheet->getColumnDimension('R')->setAutoSize(false)->setWidth(6.5);  // 整理員
        $sheet->getColumnDimension('S')->setAutoSize(false)->setWidth(35);   // 備考

        // ★ 印刷設定
        $sheet->getPageSetup()->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
        $sheet->getPageSetup()->setPaperSize(PageSetup::PAPERSIZE_A3);
        $sheet->getPageSetup()->setFitToWidth(1);
        $sheet->getPageSetup()->setFitToHeight(0); 
        
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="外業管理表_'.date('Ymd').'.xlsx"');
        header('Cache-Control: max-age=0');
        
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        exit;
    } catch (Exception $e) {
        die("Excel出力エラー: " . $e->getMessage());
    }
}


// ==========================================
// 3. 安全作業計画書 Excel生成処理 (既存)
// ==========================================
if (isset($_POST['generate_excel'])) {
    $spreadsheet = null;
    try {
        $templatePath = __DIR__ . '/template.xlsx';
        if (!file_exists($templatePath)) throw new Exception("テンプレートファイルが見つかりません。");

        $spreadsheet = IOFactory::load($templatePath);
        $sheet = $spreadsheet->getSheet(0);

        $sheet->setCellValue('L12', $_POST['job_no'] ?? '');
        $sheet->setCellValue('B12', $_POST['job_content'] ?? '');
        $sheet->setCellValue('B13', $_POST['location'] ?? '');
        $sheet->setCellValue('B35', $_POST['work_detail'] ?? '');
        $sheet->getStyle('B35')->getAlignment()->setWrapText(true);
        $sheet->setCellValue('B55', $_POST['safety_measures'] ?? '');
        $sheet->getStyle('B55')->getAlignment()->setWrapText(true);
        $sheet->setCellValue('K54', $_POST['danger_other_text'] ?? '');

        if (!empty($_POST['dangers'])) {
            $ellipsePath = __DIR__ . '/circle.png';
            if (file_exists($ellipsePath)) {
                $dangerMap = ['触車' => 'B54', '感電' => 'D54', '墜落' => 'G54', 'その他' => 'I54'];
                foreach ($_POST['dangers'] as $danger) {
                    if (isset($dangerMap[$danger])) {
                        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                        $drawing->setName('Danger_Circle');
                        $drawing->setPath($ellipsePath);
                        $drawing->setCoordinates($dangerMap[$danger]);
                        if ($danger === 'その他') {
                            $drawing->setOffsetX(115); $drawing->setOffsetY(-2);
                        } else {
                            $drawing->setOffsetX(60); $drawing->setOffsetY(-2); 
                        }
                        $drawing->setWorksheet($sheet);
                    }
                }
            }
        }

        $teamInfo = null;
        if (!empty($_POST['team_id'])) {
            $stmt = $pdo->prepare("SELECT * FROM team_settings WHERE id = ?");
            $stmt->execute([$_POST['team_id']]);
            $teamInfo = $stmt->fetch(PDO::FETCH_ASSOC);
        }

        $yorudatsu_data = [];
        if (!empty($_POST['yorudatsu_csv_data'])) {
            $yorudatsu_data = json_decode($_POST['yorudatsu_csv_data'], true) ?? [];
        }

        $has_closure_overall = false;

        for ($i = 1; $i <= 5; $i++) {
            $r_date = 6 + $i; 
            $raw_date = $_POST["date_{$i}"] ?? '';
            $sheet->setCellValue('C'.$r_date, $raw_date ? date('Y年n月j日', strtotime($raw_date)) : '');
            
            $raw_start = $_POST["start_{$i}"] ?? '';
            $raw_end = $_POST["end_{$i}"] ?? '';
            $fmt_start = $raw_start !== '' ? date('G:i', strtotime($raw_start)) : '';
            $fmt_end   = $raw_end !== '' ? date('G:i', strtotime($raw_end)) : '';
            
            $sheet->setCellValue('G'.$r_date, $fmt_start);
            $sheet->setCellValue('I'.$r_date, $fmt_end);
            $sheet->setCellValue('J'.$r_date, $_POST["reserve_{$i}"] ?? '');

            $r_our = 17 + $i;
            $sheet->setCellValue('D'.$r_our, $_POST["our_leader_{$i}"] ?? '');
            $sheet->setCellValue('E'.$r_our, $_POST["our_phone_{$i}"] ?? '');
            $sheet->setCellValue('F'.$r_our, $_POST["our_w1_{$i}"] ?? '');
            $sheet->setCellValue('G'.$r_our, $_POST["our_w2_{$i}"] ?? '');
            $sheet->setCellValue('H'.$r_our, $_POST["our_w3_{$i}"] ?? '');
            $sheet->setCellValue('I'.$r_our, $_POST["our_w4_{$i}"] ?? '');
            
            $our_cl = $_POST["our_cl_{$i}"] ?? '';
            $sheet->setCellValue('J'.$r_our, $our_cl);
            
            $sheet->setCellValue('K'.$r_our, $_POST["our_g1_{$i}"] ?? '');
            $sheet->setCellValue('L'.$r_our, $_POST["our_g2_{$i}"] ?? '');

            $r_part = 23 + $i;
            $sheet->setCellValue('D'.$r_part, $_POST["part_name_{$i}"] ?? '');
            $sheet->setCellValue('F'.$r_part, $_POST["part_leader_{$i}"] ?? '');
            $sheet->setCellValue('G'.$r_part, $_POST["part_phone_{$i}"] ?? '');
            $sheet->setCellValue('H'.$r_part, $_POST["part_count_{$i}"] ?? '');
            $sheet->setCellValue('I'.$r_part, $_POST["part_g_count_{$i}"] ?? '');
            $sheet->setCellValue('J'.$r_part, $_POST["part_t_count_{$i}"] ?? '');
            $sheet->setCellValue('K'.$r_part, $_POST["part_other_{$i}"] ?? '');

            $r_client = 29 + $i;
            $sheet->setCellValue('C'.$r_client, $_POST["client_num_{$i}"] ?? '');
            $sheet->setCellValue('D'.$r_client, $_POST["client_name_{$i}"] ?? '');

            if ($raw_date !== '') {
                if ($our_cl !== '') {
                    $has_closure_overall = true;
                    $sheetWork = $spreadsheet->getSheetByName("work{$i}d");
                    
                    if ($sheetWork !== null) {
                        if ($teamInfo) {
                            $sheetWork->setCellValue('C62', $teamInfo['group_leader_name'] ?? '');
                            $sheetWork->setCellValue('F62', $teamInfo['group_leader_phone'] ?? '');
                        }

                        $current_date = date('Y-m-d', strtotime($raw_date));
                        $next_date    = date('Y-m-d', strtotime($raw_date . ' +1 day'));
                        
                        if (isHolidayOrWeekend($current_date, $holidays)) {
                            $sheetWork->setCellValue('K13', '（土休2号）');
                        } else {
                            $sheetWork->setCellValue('K13', '（平日1号）');
                        }
                        
                        if (isHolidayOrWeekend($next_date, $holidays)) {
                            $sheetWork->setCellValue('V13', '（土休2号）');
                        } else {
                            $sheetWork->setCellValue('V13', '（平日1号）');
                        }

                        if (!empty($yorudatsu_data)) {
                            $target_date = date('Y-m-d', strtotime($raw_date));
                            $target_name = str_replace([' ', '　'], '', $our_cl); 
                            
                            foreach ($yorudatsu_data as $row) {
                                if (count($row) > 28) {
                                    $csv_date_str = trim($row[0]); 
                                    if ($csv_date_str === '') continue;
                                    $csv_date = date('Y-m-d', strtotime(str_replace('/', '-', $csv_date_str)));
                                    $csv_name = str_replace([' ', '　'], '', trim($row[8])); 
                                    
                                    if ($target_date === $csv_date && $target_name === $csv_name) {
                                        $sheetWork->setCellValue('H6', trim($row[2]));
                                        $sheetWork->setCellValue('G8', trim($row[10]));
                                        
                                        $w_val = trim($row[22]);
                                        $line_g9 = ''; $line_e22 = '';
                                        if ($w_val === '京都本線') { $line_g9 = '京都線'; $line_e22 = '京都'; }
                                        elseif ($w_val === '神戸本線') { $line_g9 = '神戸線'; $line_e22 = '神戸'; }
                                        elseif ($w_val === '宝塚本線') { $line_g9 = '宝塚線'; $line_e22 = '宝塚'; }
                                        else { $line_g9 = $w_val; $line_e22 = $w_val; }

                                        $y_val = trim($row[24]); $z_val = trim($row[25]);
                                        $place_str = '';
                                        if ($y_val !== '' || $z_val !== '') {
                                            $place_str = ($y_val === $z_val) ? $y_val . '構内' : $y_val . '～' . $z_val;
                                        }
                                        
                                        $sheetWork->setCellValue('G9', trim($line_g9 . ' ' . $place_str));
                                        $sheetWork->setCellValue('E22', $line_e22);
                                        $sheetWork->setCellValue('G10', trim($row[28]));
                                        break; 
                                    }
                                } 
                                elseif (count($row) === 8) {
                                    $csv_date_str = trim($row[0]); 
                                    if ($csv_date_str === '') continue;
                                    $csv_date = date('Y-m-d', strtotime(str_replace('/', '-', $csv_date_str)));
                                    $csv_name = str_replace([' ', '　'], '', trim($row[2])); 
                                    
                                    if ($target_date === $csv_date && $target_name === $csv_name) {
                                        $sheetWork->setCellValue('H6', trim($row[1]));
                                        $sheetWork->setCellValue('G8', trim($row[3]));
                                        
                                        $w_val = trim($row[4]);
                                        $line_g9 = ''; $line_e22 = '';
                                        if ($w_val === '京都本線') { $line_g9 = '京都線'; $line_e22 = '京都'; }
                                        elseif ($w_val === '神戸本線') { $line_g9 = '神戸線'; $line_e22 = '神戸'; }
                                        elseif ($w_val === '宝塚本線') { $line_g9 = '宝塚線'; $line_e22 = '宝塚'; }
                                        else { $line_g9 = $w_val; $line_e22 = $w_val; }

                                        $y_val = trim($row[5]); $z_val = trim($row[6]);
                                        $place_str = '';
                                        if ($y_val !== '' || $z_val !== '') {
                                            $place_str = ($y_val === $z_val) ? $y_val . '構内' : $y_val . '～' . $z_val;
                                        }
                                        
                                        $sheetWork->setCellValue('G9', trim($line_g9 . ' ' . $place_str));
                                        $sheetWork->setCellValue('E22', $line_e22);
                                        $sheetWork->setCellValue('G10', trim($row[7]));
                                        break; 
                                    }
                                }
                            }
                        }

                        $barPath = __DIR__ . '/bar600.png';
                        if (file_exists($barPath)) {
                            if (empty($_POST["chk_kido_{$i}"])) {
                                $draw = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                                $draw->setPath($barPath);
                                $draw->setCoordinates('D141'); $draw->setOffsetX(-16); $draw->setOffsetY(13); $draw->setWorksheet($sheetWork);
                            }
                            if (empty($_POST["chk_denki_{$i}"])) {
                                $draw = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                                $draw->setPath($barPath);
                                $draw->setCoordinates('D142'); $draw->setOffsetX(-16); $draw->setOffsetY(13); $draw->setWorksheet($sheetWork);
                            }
                            if (empty($_POST["chk_teiden_{$i}"])) {
                                $teidenCellsC = ['C34', 'C35', 'C36', 'C37', 'C46', 'C47'];
                                foreach ($teidenCellsC as $cell) {
                                    $draw = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                                    $draw->setPath($barPath); $draw->setCoordinates($cell); $draw->setOffsetX(5); $draw->setOffsetY(10); $draw->setWorksheet($sheetWork);
                                }
                                $teidenCellsD = ['D143'];
                                foreach ($teidenCellsD as $cell) {
                                    $draw = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                                    $draw->setPath($barPath); $draw->setCoordinates($cell); $draw->setOffsetX(-16); $draw->setOffsetY(13); $draw->setWorksheet($sheetWork);
                                }
                            }
                            if (empty($_POST["chk_toro_{$i}"])) {
                                $draw = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                                $draw->setPath($barPath); $draw->setCoordinates('D146'); $draw->setOffsetX(-16); $draw->setOffsetY(13); $draw->setWorksheet($sheetWork);
                            }
                            if (empty($_POST["chk_kanban_{$i}"])) {
                                $draw = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                                $draw->setPath($barPath); $draw->setCoordinates('D148'); $draw->setOffsetX(-16); $draw->setOffsetY(13); $draw->setWorksheet($sheetWork);
                            }
                            if (empty($_POST["chk_fumikiri_{$i}"])) {
                                $fumiCellsC = ['C45'];
                                foreach ($fumiCellsC as $cell) {
                                    $draw = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                                    $draw->setPath($barPath); $draw->setCoordinates($cell); $draw->setOffsetX(5); $draw->setOffsetY(10); $draw->setWorksheet($sheetWork);
                                }
                                $fumiCellsD = ['D151'];
                                foreach ($fumiCellsD as $cell) {
                                    $draw = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                                    $draw->setPath($barPath); $draw->setCoordinates($cell); $draw->setOffsetX(-16); $draw->setOffsetY(13); $draw->setWorksheet($sheetWork);
                                }
                            }
                            if (empty($_POST["chk_ryuchi_{$i}"])) {
                                $draw = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                                $draw->setPath($barPath); $draw->setCoordinates('C42'); $draw->setOffsetX(5); $draw->setOffsetY(10); $draw->setWorksheet($sheetWork);
                            }
                        }
                    }
                } else {
                    $sheetWork = $spreadsheet->getSheetByName("work{$i}d");
                    if ($sheetWork !== null) $spreadsheet->removeSheetByIndex($spreadsheet->getIndex($sheetWork));

                    $sheetChecklist = $spreadsheet->getSheetByName("checklist{$i}d");
                    if ($sheetChecklist !== null) {
                        $bar400Path = __DIR__ . '/bar400.png';
                        if (file_exists($bar400Path)) {
                            $offsetX_bar400 = 0; $offsetY_bar400 = 10;
                            // ★ 触車チェックの有無でbar400.pngの挿入セルを分岐
                            $is_shokusha = !empty($_POST['dangers']) && in_array('触車', $_POST['dangers']);
                            $cells400 = $is_shokusha ? ['B22', 'B30', 'B32', 'B61'] : ['B22', 'B23', 'B30', 'B32', 'B59', 'B61'];
                            
                            foreach ($cells400 as $cell) {
                                $draw = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                                $draw->setPath($bar400Path); $draw->setCoordinates($cell); $draw->setOffsetX($offsetX_bar400); $draw->setOffsetY($offsetY_bar400); $draw->setWorksheet($sheetChecklist);
                            }
                        }
                    }
                    
                    $sheetKinkyu1 = $spreadsheet->getSheetByName("緊急連絡");
                    if ($sheetKinkyu1 !== null) {
                        $baseRow = 7 + ($i - 1) * 6;
                        $sheetKinkyu1->setCellValue('Q' . $baseRow, ''); $sheetKinkyu1->setCellValue('P' . ($baseRow + 1), ''); $sheetKinkyu1->setCellValue('P' . ($baseRow + 2), '');
                    }
                }

            } else {
                $sheetChecklist = $spreadsheet->getSheetByName("checklist{$i}d");
                if ($sheetChecklist !== null) $spreadsheet->removeSheetByIndex($spreadsheet->getIndex($sheetChecklist));
                $sheetWork = $spreadsheet->getSheetByName("work{$i}d");
                if ($sheetWork !== null) $spreadsheet->removeSheetByIndex($spreadsheet->getIndex($sheetWork));
                $sheetKinkyu1 = $spreadsheet->getSheetByName("緊急連絡");
                if ($sheetKinkyu1 !== null) {
                    $baseRow = 7 + ($i - 1) * 6;
                    $sheetKinkyu1->setCellValue('Q' . $baseRow, ''); $sheetKinkyu1->setCellValue('P' . ($baseRow + 1), ''); $sheetKinkyu1->setCellValue('P' . ($baseRow + 2), '');
                }
            }
        }

        if (!$has_closure_overall) {
            $sheetKinkyu1 = $spreadsheet->getSheetByName("緊急連絡");
            if ($sheetKinkyu1 !== null) $spreadsheet->removeSheetByIndex($spreadsheet->getIndex($sheetKinkyu1));
            $sheetKinkyu2 = $spreadsheet->getSheetByName("緊急連絡 (業者様用)");
            if ($sheetKinkyu2 !== null) $spreadsheet->removeSheetByIndex($spreadsheet->getIndex($sheetKinkyu2));
        } else {
            if ($teamInfo) {
                $sheetKinkyu1 = $spreadsheet->getSheetByName("緊急連絡");
                if ($sheetKinkyu1 !== null) {
                    $sheetKinkyu1->setCellValue('Q37', $teamInfo['contact1_name']); $sheetKinkyu1->setCellValue('R37', $teamInfo['contact1_phone']);
                    $sheetKinkyu1->setCellValue('Q38', $teamInfo['contact2_name']); $sheetKinkyu1->setCellValue('R38', $teamInfo['contact2_phone']);
                }
            }
        }
        
        $spreadsheet->setActiveSheetIndex(0);
        $job_no = preg_replace('/[^a-zA-Z0-9_-]/', '', $_POST['job_no'] ?? 'draft');
        $outFileName = "安全作業計画書_" . $job_no . "_" . date('Ymd') . ".xlsx";
        
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $outFileName . '"');
        header('Cache-Control: max-age=0');
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        $spreadsheet->disconnectWorksheets(); unset($spreadsheet);
        exit;
    } catch (Exception $e) {
        if ($spreadsheet !== null) { $spreadsheet->disconnectWorksheets(); unset($spreadsheet); }
        die("<h3>実行エラー</h3><p>" . htmlspecialchars($e->getMessage(), ENT_QUOTES, 'UTF-8') . "</p>");
    }
}
?>

<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>安全作業計画書 入力・DB管理システム</title>
    <style>
        *, *::before, *::after { box-sizing: border-box; }

        body { font-family: "Meiryo", sans-serif; background-color: #f0f4f8; padding: 20px; font-size: 13px;}
        .container { max-width: 1600px; margin: auto; background: #fff; padding: 20px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
        h2 { color: #005a9e; border-bottom: 2px solid #005a9e; padding-bottom: 10px; }
        
        .toolbar { display: flex; gap: 10px; margin-bottom: 20px; background: #eef2f7; padding: 10px; border-radius: 6px; align-items: center; flex-wrap: wrap;}
        .toolbar select, .toolbar input { padding: 6px; }
        .btn { padding: 6px 12px; border-radius: 4px; border: none; cursor: pointer; font-weight: bold; color: #fff; font-size:12px;}
        .btn-blue { background: #005a9e; } .btn-green { background: #28a745; }
        .btn-yellow { background: #ffc107; color: #333; } .btn-gray { background: #6c757d; } .btn-red { background: #dc3545;}
        
        .section-title { background: #005a9e; color: #fff; padding: 8px 15px; font-weight: bold; border-radius: 4px; margin-top: 25px; margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center;}
        
        .table-wrap { overflow-x: auto; margin-bottom: 15px; border: 1px solid #ddd;}
        
        table { border-collapse: collapse; table-layout: fixed; background: #fff; }
        th, td { border: 1px solid #ddd; padding: 5px; text-align: center; vertical-align: middle; }
        th { background: #f8f9fa; font-weight: normal; font-size: 12px;}
        td input[type="text"], td input[type="time"], td input[type="date"], td select { width: 100%; padding: 5px; }
        textarea { width: 100%; padding: 10px; border: 1px solid #ccc; border-radius: 4px; resize: none; overflow: hidden; line-height: 1.5;}
        
        .day-handle { background: #ffeb3b; font-weight: bold; cursor: grab; user-select: none; border-right: 3px solid #ccc;}
        .day-handle:active { cursor: grabbing; }
        .day-handle small { display: block; font-size:10px; color:#555; }
        .drag-over { background-color: #d4edda !important; }
        .clear-btn { background: #fff; border: 1px solid #dc3545; color: #dc3545; padding: 4px; border-radius: 3px; cursor: pointer; font-size:10px; width:100%; }
        .clear-btn:hover { background: #dc3545; color: #fff; }

        .time-box { display: flex; align-items: center; justify-content: flex-start; gap: 4px;}
        .time-box button { padding: 4px; font-size: 10px; min-width: 30px;}
        
        .danger-checks label { margin-right: 15px; font-weight: bold; font-size: 14px; cursor: pointer;}
        
        .modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 100;}
        .modal-content { 
            background: #fff; 
            margin: 5vh auto; 
            padding: 20px; 
            border-radius: 8px; 
            max-height: 90vh; 
            overflow-y: auto; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        }
        .person-table { width: 100%; margin-top: 15px; font-size: 12px; table-layout: fixed;}
        .person-table th, .person-table td { padding: 6px; overflow: hidden; white-space: nowrap; text-overflow: ellipsis;}
        
        .plan-list-table th, .plan-list-table td { padding: 8px; font-size: 12px; white-space: normal; word-break: break-word; text-align: left;}
        .plan-list-table th { background: #005a9e; color:#fff; text-align:center;}
        .editable-title { border: 1px solid transparent; background: transparent; width: 100%; cursor: pointer; padding: 2px; }
        .editable-title:focus { border: 1px solid #005a9e; background: #fff; cursor: text; outline:none;}

        /* 外業管理表用スタイル */
        .gaigyo-table th, .gaigyo-table td { padding: 4px; font-size: 11px; white-space: nowrap; border: 1px solid #ccc; text-align:center; vertical-align: middle;}
        .gaigyo-table th { background: #005a9e; color:#fff; }
        .gaigyo-table input[type="text"] { width: 100%; padding: 2px; border: 1px solid #ccc; border-radius: 2px; font-size: 11px; text-align: left;}

        /* 印刷用スタイル */
        @media print {
            body { background: #fff !important; margin: 0; padding: 0; height: auto !important; overflow: visible !important; }
            .container, #dbModal, #templateModal, #teamModal, #planListModal { display: none !important; }
            #gaigyoModal { position: relative !important; display: block !important; background: transparent !important; zoom: 0.82; }
            #gaigyoModal .modal-content { box-shadow: none !important; border: none !important; width: 100% !important; max-height: none !important; overflow: visible !important; padding: 0 !important; margin: 0 !important; margin-left: 15px !important; }
            #gaigyoModal div[style*="overflow-x: auto"], #gaigyoModal div[style*="overflow-x:auto"] { overflow: visible !important; max-height: none !important; height: auto !important; }
            .btn, .toolbar { display: none !important; }
            .gaigyo-table { width: 100% !important; border-collapse: collapse !important; page-break-inside: auto; }
            .gaigyo-table tr { page-break-inside: avoid; page-break-after: auto; }
            .gaigyo-table th { background-color: #eee !important; color: #000 !important; border: 1px solid #000 !important; -webkit-print-color-adjust: exact; }
            .gaigyo-table td { border: 1px solid #000 !important; color: #000 !important; }
            .gaigyo-table input[type="text"] { border: none !important; background: transparent !important; color: #000 !important; width: 100%; }
            @page { size: A3 landscape; margin: 10mm 10mm 10mm 15mm; }
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>安全作業計画書 入力・DB管理システム</h2>
        
        <div class="toolbar">
            <button class="btn btn-blue" onclick="openDbModal()">👥 名簿管理</button>
            <span style="border-left: 2px solid #ccc; height: 20px; margin: 0 10px;"></span>
            
            <label style="font-weight:bold; font-size:12px;">DB保存名:</label>
            <input type="text" id="plan_save_name" style="width:200px;">
            <button type="button" class="btn btn-green" onclick="handleManualSave()">💾 DBに保存</button>
            
            <span style="border-left: 2px solid #ccc; height: 20px; margin: 0 10px;"></span>
            
            <select id="load_plan_select" style="max-width:250px;"><option value="">過去のデータを読み込む...</option></select>
            <button class="btn btn-gray" onclick="loadPlanFromDB()">読込</button>
            
            <button class="btn btn-yellow" onclick="openPlanListModal()" style="margin-left:5px;">📂 保存データ一覧</button>
            <button class="btn btn-blue" onclick="openGaigyoModal()" style="margin-left:5px; background-color:#17a2b8;">📋 外業管理表</button>
            <button class="btn btn-red" onclick="deletePlanFromDB()" style="margin-left:auto;">🗑️ 削除</button>
        </div>

        <form method="post" id="planForm">
            <input type="hidden" name="generate_excel" value="1">
            
            <div class="section-title">1. 基本情報</div>
            <table style="width: 100%; table-layout: auto;">
                <tr>
                    <td width="10%" style="background:#f8f9fa;">工番 (L12)</td>
                    <td width="12%" style="text-align: left;"><input type="text" name="job_no" id="input_job_no" style="width: 100px;" oninput="updateSaveName()"></td>
                    <td width="8%" style="background:#f8f9fa;">チーム</td>
                    <td width="15%">
                        <div style="display:flex; gap:5px; align-items:center;">
                            <select name="team_id" id="team_select" style="width: 120px;"><option value="">選択</option></select>
                            <button type="button" class="btn btn-gray" style="padding:4px; font-size:11px;" onclick="openTeamModal()">⚙️</button>
                        </div>
                    </td>
                    <td width="12%" style="background:#f8f9fa;">作業場所 (B13)</td>
                    <td width="43%"><input type="text" name="location"></td>
                </tr>
                <tr><td style="background:#f8f9fa;">工事内容 (B12)</td><td colspan="5"><input type="text" name="job_content"></td></tr>
                <tr><td style="background:#f8f9fa; vertical-align: top; padding-top: 10px;">作業内容 (B35)</td><td colspan="5"><textarea name="work_detail" rows="3" oninput="autoResize(this)" placeholder="複数行入力可能です。改行すると自動で広がります。"></textarea></td></tr>
            </table>

            <div class="section-title">2. 予測される危険 ＆ 安全対策</div>
            <div style="padding: 10px; border: 1px solid #ddd; background: #fff; margin-bottom: 15px;">
                <div class="danger-checks" style="margin-bottom: 10px;">
                    予測される危険 (該当箇所に赤丸を描画します): 
                    <label><input type="checkbox" name="dangers[]" value="触車"> 触車</label>
                    <label><input type="checkbox" name="dangers[]" value="感電"> 感電</label>
                    <label><input type="checkbox" name="dangers[]" value="墜落"> 墜落</label>
                    <label><input type="checkbox" name="dangers[]" value="その他" onchange="document.getElementById('danger_other_text').style.display = this.checked ? 'inline-block' : 'none';"> その他</label>
                    <input type="text" name="danger_other_text" id="danger_other_text" maxlength="10" placeholder="内容(10文字程度)" style="display:none; width:150px; margin-left:10px; padding:3px;">
                </div>
                
                <div style="display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:5px;">
                    <label>安全対策 (B55)</label>
                    <div>
                        <select name="safety_template_id" id="safety_template_select" onchange="applySafetyTemplate()" style="padding:4px; font-size:12px; min-width:225px;"><option value="">選択</option></select>
                        <button type="button" class="btn btn-gray" style="padding:4px 8px; font-size:11px;" onclick="openTemplateModal()">⚙️ テンプレ管理</button>
                    </div>
                </div>
                <textarea name="safety_measures" id="safety_measures" rows="15" oninput="autoResize(this)"><?= htmlspecialchars($default_safety) ?></textarea>
            </div>

            <div class="section-title">
                3. 作業日時
                <div style="font-size:12px; font-weight:normal; display:flex; align-items:center;">
                    <span style="background:#ffc107; color:#333; padding:2px 5px; border-radius:3px; margin-right:5px; font-weight:bold;">自動転記</span>📁 夜達CSVを取り込む: 
                    <input type="file" id="csv_input_file" accept=".csv" onclick="this.value=null;" onchange="readYorudatsuCSV(this)" style="width: auto; padding: 2px; background:#fff; color:#000; border-radius:3px; margin-left:5px;">
                    <input type="hidden" name="yorudatsu_csv_data" id="yorudatsu_csv_data">
                    <span id="csv_restore_msg" style="display:none; color:#28a745; font-weight:bold; margin-left:10px; background:#fff; padding:2px 5px; border-radius:3px;">✅ 過去のCSVデータを復元済みです</span>
                </div>
            </div>
            <div class="table-wrap">
                <table style="width: max-content;">
                    <tr>
                        <th style="width: 50px; min-width: 50px; max-width: 50px;">コピー</th>
                        <th style="width: 200px; min-width: 200px; max-width: 200px;">作業日時</th>
                        <th style="width: 280px; min-width: 280px; max-width: 280px;">時間入力</th>
                        <th style="width: 360px; min-width: 360px; max-width: 360px;">手配・立会確認</th>
                        <th style="width: 60px; min-width: 60px; max-width: 60px;">操作</th>
                    </tr>
                    <?php for($i=1; $i<=5; $i++): ?>
                    <tr>
                        <td class="day-handle" draggable="true" data-day="<?= $i ?>" ondragstart="dragStart(event)" ondragover="allowDrop(event)" ondragleave="dragLeave(event)" ondrop="dropCopy(event)">
                            <?= $i ?>日目<br><small>≡掴む≡</small>
                        </td>
                        <td>
                            <div style="display: flex; align-items: center; justify-content: center; gap: 5px;">
                                <input type="date" name="date_<?= $i ?>" style="width: auto;">
                                <label style="font-size: 11px; cursor: pointer; white-space: nowrap;"><input type="checkbox" name="reserve_<?= $i ?>" value="（予備日）"> 予備日</label>
                            </div>
                        </td>
                        <td>
                            <div class="time-box">
                                <button type="button" class="btn btn-yellow" onclick="setTime(this, '09:00', '17:00', <?= $i ?>)">昼</button>
                                <button type="button" class="btn-gray" onclick="setTime(this, '00:00', '05:00', <?= $i ?>)" style="color:#fff;">夜</button>
                                <input type="time" name="start_<?= $i ?>" style="width:90px;" onchange="syncNightLeader(<?= $i ?>)">～<input type="time" name="end_<?= $i ?>" style="width:90px;">
                            </div>
                        </td>
                        <td style="text-align: left; padding-left: 15px;">
                            <div style="display: flex; flex-wrap: wrap; gap: 6px; width: 100%;">
                                <label style="font-size: 11px; cursor: pointer; width: 75px;"><input type="checkbox" name="chk_kido_<?= $i ?>" value="1"> 軌道立会</label>
                                <label style="font-size: 11px; cursor: pointer; width: 75px;"><input type="checkbox" name="chk_denki_<?= $i ?>" value="1"> 電気立会</label>
                                <label style="font-size: 11px; cursor: pointer; width: 60px;"><input type="checkbox" name="chk_teiden_<?= $i ?>" value="1"> 停電</label>
                                <label style="font-size: 11px; cursor: pointer; width: 75px;"><input type="checkbox" name="chk_toro_<?= $i ?>" value="1"> トロ使用</label>
                                <label style="font-size: 11px; cursor: pointer; width: 85px;"><input type="checkbox" name="chk_kanban_<?= $i ?>" value="1"> 工事表示板</label>
                                <label style="font-size: 11px; cursor: pointer; width: 75px;"><input type="checkbox" name="chk_fumikiri_<?= $i ?>" value="1"> 踏切鳴止</label>
                                <label style="font-size: 11px; cursor: pointer; width: 75px;"><input type="checkbox" name="chk_ryuchi_<?= $i ?>" value="1"> 留置変更</label>
                            </div>
                        </td>
                        <td><button type="button" class="clear-btn" onclick="clearDay(this, <?= $i ?>)">クリア</button></td>
                    </tr>
                    <?php endfor; ?>
                </table>
            </div>

            <div class="section-title">
                4. 当社体制
                <div>
                    <button type="button" class="btn btn-yellow" onclick="addWorkerCol()" id="addWorkerBtn">＋ 作業員を追加表示</button>
                    <button type="button" class="btn btn-red" onclick="removeWorkerCol()" id="removeWorkerBtn" style="display:none;">－ 枠を減らす</button>
                </div>
            </div>
            <div class="table-wrap">
                <table style="width: max-content;">
                    <tr>
                        <th style="width: 50px; min-width: 50px; max-width: 50px;">コピー</th>
                        <th style="width: 140px; min-width: 140px; max-width: 140px;">当社 指揮者</th>
                        <th style="width: 140px; min-width: 140px; max-width: 140px;">携帯番号</th>
                        <th class="cw1" style="width: 140px; min-width: 140px; max-width: 140px;">作業員1</th>
                        <th class="cw2" style="width: 140px; min-width: 140px; max-width: 140px; display:none;">作業員2</th>
                        <th class="cw3" style="width: 140px; min-width: 140px; max-width: 140px; display:none;">作業員3</th>
                        <th class="cw4" style="width: 140px; min-width: 140px; max-width: 140px; display:none;">作業員4</th>
                        <th style="width: 140px; min-width: 140px; max-width: 140px;">閉鎖責任者</th>
                        <th style="width: 140px; min-width: 140px; max-width: 140px;">監視員1</th>
                        <th style="width: 140px; min-width: 140px; max-width: 140px;">監視員2</th>
                        <th style="width: 60px; min-width: 60px; max-width: 60px;">操作</th>
                    </tr>
                    <?php for($i=1; $i<=5; $i++): ?>
                    <tr>
                        <td class="day-handle" draggable="true" data-day="<?= $i ?>" ondragstart="dragStart(event)" ondragover="allowDrop(event)" ondragleave="dragLeave(event)" ondrop="dropCopy(event)">
                            <?= $i ?>日目<br><small>≡掴む≡</small>
                        </td>
                        <td style="width: 140px; min-width: 140px; max-width: 140px;"><select name="our_leader_<?= $i ?>" class="sel_our" onchange="setPhone(this, 'our_phone_<?= $i ?>'); syncNightLeader(<?= $i ?>)"><option value="">選択</option></select></td>
                        <td style="width: 140px; min-width: 140px; max-width: 140px;"><input type="text" name="our_phone_<?= $i ?>" id="our_phone_<?= $i ?>" readonly style="background:#f0f0f0;"></td>
                        <td class="cw1" style="width: 140px; min-width: 140px; max-width: 140px;"><select name="our_w1_<?= $i ?>" class="sel_our"><option value="">選択</option></select></td>
                        <td class="cw2" style="width: 140px; min-width: 140px; max-width: 140px; display:none;"><select name="our_w2_<?= $i ?>" class="sel_our"><option value="">選択</option></select></td>
                        <td class="cw3" style="width: 140px; min-width: 140px; max-width: 140px; display:none;"><select name="our_w3_<?= $i ?>" class="sel_our"><option value="">選択</option></select></td>
                        <td class="cw4" style="width: 140px; min-width: 140px; max-width: 140px; display:none;"><select name="our_w4_<?= $i ?>" class="sel_our"><option value="">選択</option></select></td>
                        <td style="width: 140px; min-width: 140px; max-width: 140px;"><select name="our_cl_<?= $i ?>" class="sel_our"><option value="">選択</option></select></td>
                        <td style="width: 140px; min-width: 140px; max-width: 140px;"><select name="our_g1_<?= $i ?>" class="sel_our"><option value="">選択</option></select></td>
                        <td style="width: 140px; min-width: 140px; max-width: 140px;"><select name="our_g2_<?= $i ?>" class="sel_our"><option value="">選択</option></select></td>
                        <td><button type="button" class="clear-btn" onclick="clearDay(this, <?= $i ?>)">クリア</button></td>
                    </tr>
                    <?php endfor; ?>
                </table>
            </div>

            <div class="section-title">5. 協力業者</div>
            <div class="table-wrap">
                <table style="width: max-content;">
                    <tr>
                        <th style="width: 50px; min-width: 50px; max-width: 50px;">コピー</th>
                        <th style="width: 250px; min-width: 250px; max-width: 250px;">協力業者 業者名</th>
                        <th style="width: 140px; min-width: 140px; max-width: 140px;">協力業者 責任者</th>
                        <th style="width: 140px; min-width: 140px; max-width: 140px;">携帯番号</th>
                        <th style="width: 60px; min-width: 60px; max-width: 60px;">従事者</th>
                        <th style="width: 60px; min-width: 60px; max-width: 60px;">監視員</th>
                        <th style="width: 60px; min-width: 60px; max-width: 60px;">誘導員</th>
                        <th style="width: 60px; min-width: 60px; max-width: 60px;">その他</th>
                        <th style="width: 60px; min-width: 60px; max-width: 60px;">操作</th>
                    </tr>
                    <?php for($i=1; $i<=5; $i++): ?>
                    <tr>
                        <td class="day-handle" draggable="true" data-day="<?= $i ?>" ondragstart="dragStart(event)" ondragover="allowDrop(event)" ondragleave="dragLeave(event)" ondrop="dropCopy(event)">
                            <?= $i ?>日目<br><small>≡掴む≡</small>
                        </td>
                        <td><select name="part_name_<?= $i ?>" class="sel_company" onchange="filterPartner(<?= $i ?>)"><option value="">選択</option></select></td>
                        <td><select name="part_leader_<?= $i ?>" id="pl_<?= $i ?>" onchange="setPhone(this, 'part_phone_<?= $i ?>')"><option value="">選択</option></select></td>
                        <td><input type="text" name="part_phone_<?= $i ?>" id="part_phone_<?= $i ?>" readonly style="background:#f0f0f0;"></td>
                        <td><input type="text" name="part_count_<?= $i ?>"></td>
                        <td><input type="text" name="part_g_count_<?= $i ?>"></td>
                        <td><input type="text" name="part_t_count_<?= $i ?>"></td>
                        <td><input type="text" name="part_other_<?= $i ?>"></td>
                        <td><button type="button" class="clear-btn" onclick="clearDay(this, <?= $i ?>)">クリア</button></td>
                    </tr>
                    <?php endfor; ?>
                </table>
            </div>

            <div class="section-title">6. 発注者立会人</div>
            <div class="table-wrap">
                <table style="width: max-content;">
                    <tr>
                        <th style="width: 50px; min-width: 50px; max-width: 50px;">コピー</th>
                        <th style="width: 100px; min-width: 100px; max-width: 100px;">人数 (C列)</th>
                        <th style="width: 500px; min-width: 500px; max-width: 500px;">所属部署・氏名 (D列)</th>
                        <th style="width: 60px; min-width: 60px; max-width: 60px;">操作</th>
                    </tr>
                    <?php for($i=1; $i<=5; $i++): ?>
                    <tr>
                        <td class="day-handle" draggable="true" data-day="<?= $i ?>" ondragstart="dragStart(event)" ondragover="allowDrop(event)" ondragleave="dragLeave(event)" ondrop="dropCopy(event)">
                            <?= $i ?>日目<br><small>≡掴む≡</small>
                        </td>
                        <td><input type="text" name="client_num_<?= $i ?>"></td>
                        <td><input type="text" name="client_name_<?= $i ?>"></td>
                        <td><button type="button" class="clear-btn" onclick="clearDay(this, <?= $i ?>)">クリア</button></td>
                    </tr>
                    <?php endfor; ?>
                </table>
            </div>

            <button type="submit" class="btn btn-blue" style="width: 100%; padding: 20px; font-size: 20px; margin-top:20px;">Excel を生成してダウンロード</button>
        </form>
    </div>

    <div id="gaigyoModal" class="modal">
        <div class="modal-content" style="max-width: 98%; width: 1800px;">
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom: 2px solid #ccc; padding-bottom:5px; margin-bottom:10px;">
                <h3 style="margin:0;">📋 外業管理表 <small style="color:#666; font-weight:normal; margin-left:10px;">(※テキストボックスは自由に追記できます)</small></h3>
                <div>
                    <button class="btn btn-gray" onclick="window.print()" style="margin-right:5px;">🖨️ 印刷する</button>
                    <button class="btn btn-green" onclick="exportGaigyoExcel()" style="margin-right:15px;">📊 Excel出力</button>
                    <button class="btn btn-red" onclick="document.getElementById('gaigyoModal').style.display='none'">閉じる</button>
                </div>
            </div>
            
            <div style="overflow-x: auto; max-height: 75vh;">
                <table class="gaigyo-table" id="gaigyoTable" style="width: 100%; border-collapse: collapse;">
                    <thead>
                        <tr>
                            <th style="width:30px;">No.</th>
                            <th style="width:90px;">作業日</th>
                            <th style="width:40px;">曜日</th>
                            <th style="width:150px;">業務名</th>
                            <th style="width:80px;">作業指揮者</th>
                            <th style="width:110px;">携帯番号</th>
                            <th style="width:120px;">作業員</th>
                            <th style="width:50px;">昼夜別</th>
                            <th style="width:90px;">時間帯</th>
                            <th style="width:80px;">夜達番号等</th>
                            <th style="width:120px;">関連夜達留変等</th>
                            <th style="width:120px;">場所</th>
                            <th style="width:120px;">業者名①</th>
                            <th style="width:80px;">作業責任者</th>
                            <th style="width:110px;">業者携帯</th>
                            <th style="width:40px;">人数</th>
                            <th style="width:40px;">列監</th>
                            <th style="width:40px;">整理員</th>
                            <th style="width:120px;">備考</th>
                        </tr>
                    </thead>
                    <tbody>
                        </tbody>
                </table>
            </div>
        </div>
    </div>

    <div id="planListModal" class="modal">
        <div class="modal-content" style="max-width: 95%; width: 1400px;">
            <h3 style="margin-top:0; border-bottom: 2px solid #ccc; padding-bottom:5px;">📂 保存データ一覧</h3>
            <div style="overflow-x: auto; max-height: 70vh;">
                <table class="plan-list-table" id="planListTableDetail" style="width: 100%; border-collapse: collapse; min-width: 1200px;">
                    <thead>
                        <tr>
                            <th style="width:150px;">DB保存名 <br><small>(クリックで名前変更)</small></th>
                            <th style="width:80px;">工番</th>
                            <th style="width:80px;">チーム</th>
                            <th style="width:150px;">作業場所</th>
                            <th style="width:200px;">工事内容</th>
                            <th style="width:120px;">手配・立会確認</th>
                            <th style="width:120px;">作業日時</th>
                            <th style="width:150px;">外注業者</th>
                            <th style="width:120px;">操作</th>
                        </tr>
                    </thead>
                    <tbody>
                        </tbody>
                </table>
            </div>
            <div style="text-align:right; margin-top:15px;">
                <button class="btn btn-gray" onclick="document.getElementById('planListModal').style.display='none'">閉じる</button>
            </div>
        </div>
    </div>

    <div id="dbModal" class="modal">
        <div class="modal-content" style="max-width: 600px;">
            <h3 style="margin-top:0; border-bottom: 2px solid #ccc; padding-bottom:5px;">名簿の登録と管理</h3>
            <div style="background:#f9f9f9; padding:15px; border-radius:5px; margin-bottom:15px;">
                <b>1件ずつ登録</b><br><br>
                <label>区分:</label> <select id="db_type"><option value="our">当社</option><option value="partner">協力業者</option></select><br><br>
                <label>業者名 (当社の場合は空白可):</label> <input type="text" id="db_company" style="width:100%;"><br><br>
                <label>氏名 (責任者名など):</label> <input type="text" id="db_name" style="width:100%;"><br><br>
                <label>携帯番号:</label> <input type="text" id="db_phone" style="width:100%;"><br><br>
                <button class="btn btn-green" onclick="savePerson()">＋ 登録する</button>
            </div>
            <div style="background:#eef2f7; padding:15px; border-radius:5px; margin-bottom:20px;">
                <b>CSV一括登録</b> <small style="color:#555;">(フォーマット: 区分[our/partner], 業者名, 氏名, 携帯番号)</small><br><br>
                <input type="file" id="csv_file" accept=".csv" style="margin-bottom: 10px;"><br>
                <button class="btn btn-blue" onclick="importCSV()">📁 CSVを読み込んで登録</button>
            </div>
            <b>登録済み名簿一覧</b>
            <table class="person-table" id="personListTable">
                <tr>
                    <th style="width: 60px;">区分</th>
                    <th style="width: auto;">業者名</th>
                    <th style="width: 100px;">氏名</th>
                    <th style="width: 110px;">携帯番号</th>
                    <th style="width: 60px;">操作</th>
                </tr>
            </table>
            <div style="text-align:right; margin-top:15px;">
                <button class="btn btn-gray" onclick="document.getElementById('dbModal').style.display='none'">閉じる</button>
            </div>
        </div>
    </div>

    <div id="templateModal" class="modal">
        <div class="modal-content" style="max-width: 900px;">
            <h3 style="margin-top:0; border-bottom: 2px solid #ccc; padding-bottom:5px;">安全対策 テンプレート管理</h3>
            <div style="background:#f9f9f9; padding:15px; border-radius:5px; margin-bottom:15px;">
                <b id="template_form_title">新規作成</b><br><br>
                <input type="hidden" id="tmpl_id">
                <label style="font-weight:bold; display:block; margin-bottom:5px;">タイトル:</label>
                <input type="text" id="tmpl_title" style="width: 250px; margin-bottom:10px;" placeholder="例：軌道内昼間作業用"><br>
                <label style="font-weight:bold; display:block; margin-bottom:5px;">安全対策 本文:</label> 
                <textarea id="tmpl_content" rows="6" style="width:100%;" placeholder="ここに安全対策の文章を入力してください"></textarea><br><br>
                <button class="btn btn-green" onclick="saveTemplate()" id="tmpl_save_btn">＋ 登録する</button>
                <button class="btn btn-gray" onclick="resetTemplateForm()" style="margin-left:10px;">クリア</button>
            </div>
            <b>登録済みテンプレート一覧</b>
            <table class="person-table" id="templateListTable">
                <tr>
                    <th style="width: auto;">タイトル</th>
                    <th style="width: 100px;">操作</th>
                </tr>
            </table>
            <div style="text-align:right; margin-top:15px;">
                <button class="btn btn-gray" onclick="document.getElementById('templateModal').style.display='none'">閉じる</button>
            </div>
        </div>
    </div>

    <div id="teamModal" class="modal">
        <div class="modal-content" style="max-width: 850px;">
            <h3 style="margin-top:0; border-bottom: 2px solid #ccc; padding-bottom:5px;">チーム設定 (緊急連絡用) 管理</h3>
            
            <div style="background:#f9f9f9; padding:15px; border-radius:5px; margin-bottom:15px;">
                <b id="team_form_title">新規チーム追加</b><br><br>
                <input type="hidden" id="team_id_input">
                
                <div style="border-bottom: 1px dashed #ccc; padding-bottom: 10px; margin-bottom: 10px;">
                    <b style="color:#005a9e; font-size:11px;">【グループ情報】 (別紙転記用)</b><br>
                    <div style="display: flex; gap: 20px; align-items: flex-end; margin-top: 5px;">
                        <div>
                            <label style="font-size:11px; display:block; margin-bottom:3px;">グループ名</label>
                            <input type="text" id="team_group_name" style="width: 150px;">
                        </div>
                        <div>
                            <label style="font-size:11px; display:block; margin-bottom:3px;">グループ長 氏名 (C62)</label>
                            <input type="text" id="team_leader_name" style="width: 130px;">
                        </div>
                        <div>
                            <label style="font-size:11px; display:block; margin-bottom:3px;">グループ長 電話 (F62)</label>
                            <input type="text" id="team_leader_phone" style="width: 130px;">
                        </div>
                    </div>
                </div>

                <div>
                    <b style="color:#005a9e; font-size:11px;">【チーム・連絡先情報】 (メインシート転記用)</b><br>
                    <div style="display: flex; gap: 20px; align-items: flex-start; margin-top: 5px;">
                        <div>
                            <label style="font-size:11px; display:block; margin-bottom:3px;">チーム名</label>
                            <input type="text" id="team_name" style="width: 150px;">
                        </div>
                        <div>
                            <label style="font-size:11px; display:block; margin-bottom:3px;">連絡先1 氏名 (Q37)</label>
                            <input type="text" id="team_c1_name" style="width: 130px; margin-bottom:5px;"><br>
                            <label style="font-size:11px; display:block; margin-bottom:3px;">連絡先1 電話 (R37)</label>
                            <input type="text" id="team_c1_phone" style="width: 130px;">
                        </div>
                        <div>
                            <label style="font-size:11px; display:block; margin-bottom:3px;">連絡先2 氏名 (Q38)</label>
                            <input type="text" id="team_c2_name" style="width: 130px; margin-bottom:5px;"><br>
                            <label style="font-size:11px; display:block; margin-bottom:3px;">連絡先2 電話 (R38)</label>
                            <input type="text" id="team_c2_phone" style="width: 130px;">
                        </div>
                    </div>
                </div>
                
                <div style="margin-top: 15px; text-align: right;">
                    <button class="btn btn-gray" onclick="resetTeamForm()" style="margin-right:10px;">クリア</button>
                    <button class="btn btn-green" onclick="saveTeam()" id="team_save_btn">＋ 登録する</button>
                </div>
            </div>
            
            <b>登録済みチーム一覧</b>
            <table class="person-table" id="teamListTable">
                <colgroup>
                    <col style="width: 150px;">
                    <col style="width: 120px;">
                    <col style="width: 140px;">
                    <col style="width: 140px;">
                    <col style="width: 100px;">
                </colgroup>
                <tr><th>グループ情報</th><th>チーム名</th><th>連絡先1</th><th>連絡先2</th><th>操作</th></tr>
            </table>
            <div style="text-align:right; margin-top:15px;">
                <button class="btn btn-gray" onclick="document.getElementById('teamModal').style.display='none'">閉じる</button>
            </div>
        </div>
    </div>

    <script>
        let masterData = [];
        let workerCols = 1;
        let safetyTemplates = [];
        let teamSettings = [];
        
        let loadedPlanId = null; 
        let loadedDates = ['', '', '', '', ''];
        let loadedTitle = '';

        window.onload = function() {
            autoResize(document.getElementsByName('work_detail')[0]);
            autoResize(document.getElementsByName('safety_measures')[0]);
            loadMasterData();
            loadPlanList();
            loadSafetyTemplates();
            loadTeams();
            updateWorkerButtons();
            updateSaveName();
        };

        function formatTime(t) {
            if(!t) return '';
            let parts = t.split(':');
            if(parts.length === 2) return parseInt(parts[0], 10) + ':' + parts[1];
            return t;
        }

        // ==========================================
        // ★ Excel生成時のDB自動保存＆確認ロジック
        // ==========================================
        document.getElementById('planForm').addEventListener('submit', function(e) {
            e.preventDefault(); 

            const saveName = document.getElementById('plan_save_name').value || '未定';
            
            if (confirm(`“${saveName}” の名称でデータベースに保存しますか？\n(「いいえ」を選ぶと保存せずにExcelのみ生成します)`)) {
                
                let currentDates = [];
                for(let i=1; i<=5; i++) currentDates.push(document.getElementsByName('date_'+i)[0].value);
                
                if (loadedPlanId && JSON.stringify(loadedDates) === JSON.stringify(currentDates) && loadedTitle === saveName) {
                    if (confirm("⚠️ 日付も保存名も変更されていないため、既存のデータを「上書き保存」します。\nよろしいですか？\n(「キャンセル」を押すと別のデータとして新規保存します)")) {
                        savePlanToDB(true, () => { HTMLFormElement.prototype.submit.call(document.getElementById('planForm')); });
                    } else {
                        savePlanToDB(false, () => { HTMLFormElement.prototype.submit.call(document.getElementById('planForm')); });
                    }
                } else {
                    savePlanToDB(false, () => { HTMLFormElement.prototype.submit.call(document.getElementById('planForm')); });
                }
            } else {
                HTMLFormElement.prototype.submit.call(document.getElementById('planForm'));
            }
        });

        // ==========================================
        // ★ 外業管理表モーダルの処理
        // ==========================================
        function openGaigyoModal() {
            document.getElementById('gaigyoModal').style.display = 'block';
            renderGaigyoTable();
        }

        function renderGaigyoTable() {
            const req = new FormData();
            req.append('ajax_action', 'get_plans_all');
            fetch('', { method: 'POST', body: req })
                .then(r => r.json())
                .then(data => {
                    let rows = [];
                    data.forEach(p => {
                        let parsed = {};
                        try { parsed = JSON.parse(p.form_data); } catch(e) {}
                        
                        const jobContent = parsed.job_content || '';
                        const location = parsed.location || '';
                        
                        let yorudatsuData = [];
                        if(parsed.yorudatsu_csv_data) {
                            try { yorudatsuData = JSON.parse(parsed.yorudatsu_csv_data); } catch(e) {}
                        }

                        for(let i=1; i<=5; i++) {
                            const rawDate = parsed['date_'+i];
                            if(!rawDate) continue;
                            
                            const reserve = parsed['reserve_'+i] ? ' （予備日）' : '';
                            const gyomuName = jobContent + reserve;
                            
                            const leader = parsed['our_leader_'+i] || '';
                            const phone = parsed['our_phone_'+i] || '';
                            
                            let workers = [];
                            for(let w=1; w<=4; w++) {
                                if(parsed['our_w'+w+'_'+i]) workers.push(parsed['our_w'+w+'_'+i]);
                            }
                            const workerStr = workers.join(', ');
                            
                            const start = parsed['start_'+i] || '';
                            const end = parsed['end_'+i] || '';
                            let dayNight = '';
                            if(start) {
                                const h = parseInt(start.split(':')[0], 10);
                                dayNight = (h >= 6 && h < 18) ? '昼' : '夜';
                            }
                            const timeStr = (start || end) ? `${formatTime(start)}～${formatTime(end)}` : '';
                            
                            let yoruNo = '';
                            const ourCl = parsed['our_cl_'+i] || '';
                            if(ourCl && yorudatsuData.length > 0) {
                                const targetDate = new Date(rawDate).toISOString().split('T')[0];
                                const targetName = ourCl.replace(/[\s　]/g, '');
                                for(let r of yorudatsuData) {
                                    if(r.length > 28 && r[0].trim() !== '') {
                                        let cDateStr = r[0].trim().replace(/\//g, '-');
                                        let dParts = cDateStr.split('-');
                                        if(dParts.length === 3) cDateStr = `${dParts[0]}-${dParts[1].padStart(2,'0')}-${dParts[2].padStart(2,'0')}`;
                                        const cName = r[8] ? r[8].replace(/[\s　]/g, '') : '';
                                        if(targetDate === cDateStr && targetName === cName) { yoruNo = r[2] ? r[2].trim() : ''; break; }
                                    } else if (r.length === 8 && r[0].trim() !== '') {
                                        let cDateStr = r[0].trim().replace(/\//g, '-');
                                        let dParts = cDateStr.split('-');
                                        if(dParts.length === 3) cDateStr = `${dParts[0]}-${dParts[1].padStart(2,'0')}-${dParts[2].padStart(2,'0')}`;
                                        const cName = r[2] ? r[2].replace(/[\s　]/g, '') : '';
                                        if(targetDate === cDateStr && targetName === cName) { yoruNo = r[1] ? r[1].trim() : ''; break; }
                                    }
                                }
                            }
                            
                            if (yoruNo) {
                                let numMatch = yoruNo.match(/\d+/);
                                if (numMatch) { yoruNo = yoruNo.replace(numMatch[0], String(parseInt(numMatch[0], 10)).padStart(3, '0')); }
                            }
                            
                            const isRyuchi = parsed['chk_ryuchi_'+i] ? '留置変更あり' : '';
                            const partName = parsed['part_name_'+i] || '';
                            const partLeader = parsed['part_leader_'+i] || '';
                            const partPhone = parsed['part_phone_'+i] || '';
                            const partCount = parsed['part_count_'+i] || '';
                            const partGCount = parsed['part_g_count_'+i] || '';
                            const partTCount = parsed['part_t_count_'+i] || '';
                            const partOther = parsed['part_other_'+i] || '';

                            rows.push({
                                rawDate: rawDate, gyomuName: gyomuName, leader: leader, phone: phone, workerStr: workerStr,
                                dayNight: dayNight, timeStr: timeStr, yoruNo: yoruNo, isRyuchi: isRyuchi, location: location,
                                partName: partName, partLeader: partLeader, partPhone: partPhone, partCount: partCount,
                                partGCount: partGCount, partTCount: partTCount, partOther: partOther
                            });
                        }
                    });
                    
                    rows.sort((a, b) => new Date(a.rawDate) - new Date(b.rawDate));
                    
                    const tbody = document.querySelector('#gaigyoTable tbody');
                    tbody.innerHTML = '';
                    const dayOfWeekStr = ['日', '月', '火', '水', '木', '金', '土'];
                    
                    rows.forEach((r, idx) => {
                        const d = new Date(r.rawDate);
                        const w = dayOfWeekStr[d.getDay()];
                        const dateFmt = `${d.getFullYear()}/${d.getMonth()+1}/${d.getDate()}`;
                        
                        const tr = document.createElement('tr');
                        let displayGyomuName = r.gyomuName.replace('（予備日）', '<span style="color:#dc3545; font-size:10px; border:1px solid #dc3545; padding:1px 3px; border-radius:2px; margin-left:3px;">（予備日）</span>');

                        tr.innerHTML = `
                            <td>${idx + 1}</td>
                            <td>${dateFmt}</td>
                            <td>${w}</td>
                            <td style="text-align:left; white-space:normal; min-width:150px;">${displayGyomuName}</td>
                            <td>${r.leader}</td>
                            <td style="text-align:center;">${r.phone}</td>
                            <td style="text-align:left; white-space:normal; min-width:120px;">${r.workerStr}</td>
                            <td style="font-weight:bold; text-align:center;">${r.dayNight}</td>
                            <td style="text-align:center;">${r.timeStr}</td>
                            <td style="font-weight:bold; color:#005a9e; font-size:12px; text-align:center;">${r.yoruNo}</td>
                            <td><input type="text" value="${r.isRyuchi}" placeholder="入力可"></td>
                            <td style="text-align:left; white-space:normal; min-width:120px;">${r.location}</td>
                            <td style="text-align:left; white-space:normal;">${r.partName}</td>
                            <td>${r.partLeader}</td>
                            <td style="text-align:center;">${r.partPhone}</td>
                            <td style="text-align:center;">${r.partCount}</td>
                            <td style="text-align:center;">${r.partGCount}</td>
                            <td style="text-align:center;">${r.partTCount}</td>
                            <td><input type="text" value="${r.partOther}" placeholder="入力可"></td>
                        `;
                        tr.dataset.rawGyomu = r.gyomuName;
                        tbody.appendChild(tr);
                    });
                });
        }

        function exportGaigyoExcel() {
            let table = document.getElementById('gaigyoTable');
            let rows = table.querySelectorAll('tbody tr');
            let data = [];
            
            rows.forEach(tr => {
                let rowData = [];
                let tds = tr.querySelectorAll('td');
                tds.forEach((td, idx) => {
                    let input = td.querySelector('input');
                    if(input) {
                        rowData.push(input.value); 
                    } else if (idx === 3 && tr.dataset.rawGyomu) {
                        rowData.push(tr.dataset.rawGyomu);
                    } else {
                        rowData.push(td.innerText.replace(/\n/g, ' ').trim());
                    }
                });
                data.push(rowData);
            });
            
            let form = document.createElement('form');
            form.method = 'POST';
            form.action = '';
            
            let inputAction = document.createElement('input');
            inputAction.type = 'hidden';
            inputAction.name = 'export_gaigyo_excel';
            inputAction.value = '1';
            form.appendChild(inputAction);
            
            let inputData = document.createElement('input');
            inputData.type = 'hidden';
            inputData.name = 'gaigyo_data';
            inputData.value = JSON.stringify(data);
            form.appendChild(inputData);
            
            document.body.appendChild(form);
            form.submit();
            document.body.removeChild(form);
        }

        // ==========================================
        // 保存データ一覧モーダルの処理
        // ==========================================
        function openPlanListModal() {
            document.getElementById('planListModal').style.display = 'block';
            renderPlanListTable();
        }
        
        function renderPlanListTable() {
            const req = new FormData();
            req.append('ajax_action', 'get_plans_all');
            fetch('', { method: 'POST', body: req })
                .then(r => r.json())
                .then(data => {
                    const tbody = document.querySelector('#planListTableDetail tbody');
                    tbody.innerHTML = '';
                    
                    data.forEach(p => {
                        let parsed = {};
                        try { parsed = JSON.parse(p.form_data); } catch(e) {}
                        
                        const jobNo = parsed.job_no || '';
                        let teamName = '';
                        if(parsed.team_id) {
                            const t = teamSettings.find(x => x.id == parsed.team_id);
                            if(t) teamName = t.team_name;
                        }
                        const loc = parsed.location || '';
                        const content = parsed.job_content || '';
                        
                        let dates = [];
                        for(let i=1; i<=5; i++) {
                            if(parsed['date_'+i]) {
                                let d = parsed['date_'+i].split('-');
                                dates.push(`${parseInt(d[1])}/${parseInt(d[2])}`);
                            }
                        }
                        const datesStr = dates.join(', ');
                        
                        let parts = new Set();
                        for(let i=1; i<=5; i++) {
                            if(parsed['part_name_'+i]) parts.add(parsed['part_name_'+i]);
                        }
                        const partsStr = Array.from(parts).join('、');
                        
                        let checks = new Set();
                        for(let i=1; i<=5; i++) {
                            if(parsed['chk_kido_'+i]) checks.add('軌道');
                            if(parsed['chk_denki_'+i]) checks.add('電気');
                            if(parsed['chk_teiden_'+i]) checks.add('停電');
                            if(parsed['chk_toro_'+i]) checks.add('トロ');
                            if(parsed['chk_kanban_'+i]) checks.add('看板');
                            if(parsed['chk_fumikiri_'+i]) checks.add('踏切');
                            if(parsed['chk_ryuchi_'+i]) checks.add('留置');
                        }
                        const checksStr = Array.from(checks).join('、');
                        
                        const tr = document.createElement('tr');
                        tr.style.borderBottom = '1px solid #eee';
                        tr.innerHTML = `
                            <td>
                                <input type="text" class="editable-title" value="${p.title}" onchange="updatePlanTitle(${p.id}, this.value)" title="クリックして名前を変更">
                                <br><small style="color:#777;">${p.created_at.split(' ')[0]}</small>
                            </td>
                            <td>${jobNo}</td>
                            <td>${teamName}</td>
                            <td>${loc}</td>
                            <td><div style="max-height:40px; overflow:hidden; text-overflow:ellipsis;">${content}</div></td>
                            <td style="color:#d32f2f; font-weight:bold;">${checksStr}</td>
                            <td style="font-weight:bold;">${datesStr}</td>
                            <td>${partsStr}</td>
                            <td>
                                <button class="btn btn-blue" style="padding:4px 8px; margin-bottom:3px; width:100%;" onclick="loadPlanDirect(${p.id})">読込</button>
                                <button class="btn btn-red" style="padding:4px 8px; width:100%;" onclick="deletePlanDirect(${p.id})">削除</button>
                            </td>
                        `;
                        tbody.appendChild(tr);
                    });
                });
        }
        
        function updatePlanTitle(id, newTitle) {
            if(!newTitle.trim()) { alert('タイトルを入力してください。'); return; }
            const req = new FormData();
            req.append('ajax_action', 'update_plan_title');
            req.append('id', id);
            req.append('title', newTitle.trim());
            fetch('', { method: 'POST', body: req })
                .then(r => r.json())
                .then(res => { if(res.status === 'success') loadPlanList(); });
        }

        function loadPlanDirect(id) {
            document.getElementById('load_plan_select').value = id;
            loadPlanFromDB();
            document.getElementById('planListModal').style.display = 'none';
        }

        function deletePlanDirect(id) {
            if(!confirm('この保存データを削除しますか？')) return;
            const req = new FormData(); req.append('ajax_action', 'delete_plan'); req.append('id', id);
            fetch('', { method: 'POST', body: req }).then(r => r.json()).then(res => { if(res.status === 'success') { renderPlanListTable(); loadPlanList(); }});
        }

        // ==========================================
        // 既存の基本機能群
        // ==========================================
        function autoResize(el) { el.style.height = 'auto'; el.style.height = (el.scrollHeight) + 'px'; }

        function readYorudatsuCSV(input) {
            if(!input.files.length) {
                document.getElementById('yorudatsu_csv_data').value = '';
                document.getElementById('csv_restore_msg').style.display = 'none';
                return;
            }
            const file = input.files[0];
            const reader = new FileReader();
            reader.onload = function(e) {
                const text = e.target.result;
                const rows = parseCSV(text);
                
                const lightRows = rows.map(r => {
                    if (r.length > 28) return [ r[0], r[2], r[8], r[10], r[22], r[24], r[25], r[28] ];
                    return null;
                }).filter(r => r !== null);

                document.getElementById('yorudatsu_csv_data').value = JSON.stringify(lightRows);
                document.getElementById('csv_restore_msg').style.display = 'none';
                alert('✅ 夜達CSVの読み込みが完了しました！\nExcel生成時に該当日の別紙(workシート)へ自動転記されます。');
            };
            reader.readAsText(file, 'Shift_JIS');
        }

        function parseCSV(text) {
            let result = []; let row = []; let inQuotes = false; let field = '';
            for (let i = 0; i < text.length; i++) {
                let char = text[i];
                if (inQuotes) {
                    if (char === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; } } else { field += char; }
                } else {
                    if (char === '"') { inQuotes = true; } else if (char === ',') { row.push(field); field = ''; } else if (char === '\n') { row.push(field); result.push(row); row = []; field = ''; } else if (char !== '\r') { field += char; }
                }
            }
            if (field || row.length > 0) { row.push(field); result.push(row); }
            return result;
        }

        function loadTeams() {
            const formData = new FormData(); formData.append('ajax_action', 'get_teams');
            fetch('', { method: 'POST', body: formData }).then(r => r.json()).then(data => {
                teamSettings = data; const sel = document.getElementById('team_select'); const currentVal = sel.value; sel.innerHTML = '<option value="">選択</option>';
                data.forEach(t => { const opt = document.createElement('option'); opt.value = t.id; opt.text = t.team_name; sel.appendChild(opt); }); sel.value = currentVal; renderTeamTable();
            });
        }

        function openTeamModal() { document.getElementById('teamModal').style.display='block'; resetTeamForm(); }

        function resetTeamForm() {
            document.getElementById('team_id_input').value = ''; document.getElementById('team_group_name').value = ''; document.getElementById('team_leader_name').value = '';
            document.getElementById('team_leader_phone').value = ''; document.getElementById('team_name').value = ''; document.getElementById('team_c1_name').value = '';
            document.getElementById('team_c1_phone').value = ''; document.getElementById('team_c2_name').value = ''; document.getElementById('team_c2_phone').value = '';
            document.getElementById('team_form_title').innerText = '新規チーム追加'; document.getElementById('team_save_btn').innerText = '＋ 登録する'; document.getElementById('team_save_btn').className = 'btn btn-green';
        }

        function saveTeam() {
            const id = document.getElementById('team_id_input').value; const name = document.getElementById('team_name').value.trim(); if(!name) { alert('チーム名を入力してください。'); return; }
            const formData = new FormData(); formData.append('ajax_action', 'save_team'); formData.append('id', id); formData.append('group_name', document.getElementById('team_group_name').value.trim()); formData.append('group_leader_name', document.getElementById('team_leader_name').value.trim()); formData.append('group_leader_phone', document.getElementById('team_leader_phone').value.trim()); formData.append('team_name', name); formData.append('contact1_name', document.getElementById('team_c1_name').value.trim()); formData.append('contact1_phone', document.getElementById('team_c1_phone').value.trim()); formData.append('contact2_name', document.getElementById('team_c2_name').value.trim()); formData.append('contact2_phone', document.getElementById('team_c2_phone').value.trim());
            fetch('', { method: 'POST', body: formData }).then(r => r.json()).then(res => { if(res.status === 'success') { resetTeamForm(); loadTeams(); }});
        }

        function editTeam(id) {
            const t = teamSettings.find(x => x.id == id);
            if(t) { document.getElementById('team_id_input').value = t.id; document.getElementById('team_group_name').value = t.group_name; document.getElementById('team_leader_name').value = t.group_leader_name; document.getElementById('team_leader_phone').value = t.group_leader_phone; document.getElementById('team_name').value = t.team_name; document.getElementById('team_c1_name').value = t.contact1_name; document.getElementById('team_c1_phone').value = t.contact1_phone; document.getElementById('team_c2_name').value = t.contact2_name; document.getElementById('team_c2_phone').value = t.contact2_phone; document.getElementById('team_form_title').innerText = 'チームの編集'; document.getElementById('team_save_btn').innerText = '💾 更新する'; document.getElementById('team_save_btn').className = 'btn btn-blue'; }
        }

        function deleteTeam(id) {
            if(!confirm('このチームを削除しますか？')) return;
            const formData = new FormData(); formData.append('ajax_action', 'delete_team'); formData.append('id', id);
            fetch('', { method: 'POST', body: formData }).then(r => r.json()).then(res => { if(res.status === 'success') loadTeams(); });
        }

        function renderTeamTable() {
            const tbody = document.getElementById('teamListTable'); tbody.innerHTML = '<tr><th>グループ情報</th><th>チーム名</th><th>連絡先1</th><th>連絡先2</th><th>操作</th></tr>';
            teamSettings.forEach(t => { tbody.innerHTML += `<tr><td style="white-space:normal;"><b>${t.group_name}</b><br><small>${t.group_leader_name}</small><br><small>${t.group_leader_phone}</small></td><td style="font-weight:bold; white-space:normal;">${t.team_name}</td><td style="white-space:normal;">${t.contact1_name}<br><small>${t.contact1_phone}</small></td><td style="white-space:normal;">${t.contact2_name}<br><small>${t.contact2_phone}</small></td><td><button class="btn btn-blue" style="padding:2px 5px;" onclick="editTeam(${t.id})">編集</button> <button class="btn btn-red" style="padding:2px 5px;" onclick="deleteTeam(${t.id})">削除</button></td></tr>`; });
        }

        function loadSafetyTemplates() {
            const formData = new FormData(); formData.append('ajax_action', 'get_templates');
            fetch('', { method: 'POST', body: formData }).then(r => r.json()).then(data => {
                safetyTemplates = data; const sel = document.getElementById('safety_template_select'); const currentVal = sel.value; sel.innerHTML = '<option value="">選択</option>';
                data.forEach(t => { const opt = document.createElement('option'); opt.value = t.id; opt.text = t.title; sel.appendChild(opt); }); if (currentVal) sel.value = currentVal; renderTemplateTable();
            });
        }

        function applySafetyTemplate() {
            const id = document.getElementById('safety_template_select').value; if(!id) return;
            const t = safetyTemplates.find(x => x.id == id);
            if(t) { const el = document.getElementById('safety_measures'); if (el.value.trim() !== '' && el.value.trim() !== t.content.trim() && !confirm(`現在の入力内容を「${t.title}」で上書きしますか？`)) return; el.value = t.content; autoResize(el); }
        }

        function openTemplateModal() { document.getElementById('templateModal').style.display='block'; resetTemplateForm(); }

        function resetTemplateForm() {
            document.getElementById('tmpl_id').value = ''; document.getElementById('tmpl_title').value = ''; document.getElementById('tmpl_content').value = '';
            document.getElementById('template_form_title').innerText = '新規作成'; document.getElementById('tmpl_save_btn').innerText = '＋ 登録する'; document.getElementById('tmpl_save_btn').className = 'btn btn-green';
        }

        function saveTemplate() {
            const id = document.getElementById('tmpl_id').value; const title = document.getElementById('tmpl_title').value.trim(); const content = document.getElementById('tmpl_content').value.trim();
            if(!title || !content) { alert('タイトルと本文を入力してください。'); return; }
            const formData = new FormData(); formData.append('ajax_action', 'save_template'); formData.append('id', id); formData.append('title', title); formData.append('content', content);
            fetch('', { method: 'POST', body: formData }).then(r => r.json()).then(res => { if(res.status === 'success') { resetTemplateForm(); loadSafetyTemplates(); } });
        }

        function editTemplate(id) {
            const t = safetyTemplates.find(x => x.id == id);
            if(t) { document.getElementById('tmpl_id').value = t.id; document.getElementById('tmpl_title').value = t.title; document.getElementById('tmpl_content').value = t.content; document.getElementById('template_form_title').innerText = 'テンプレートの編集'; document.getElementById('tmpl_save_btn').innerText = '💾 更新する'; document.getElementById('tmpl_save_btn').className = 'btn btn-blue'; }
        }

        function deleteTemplate(id) {
            if(!confirm('このテンプレートを削除しますか？\n（元に戻すことはできません）')) return;
            const formData = new FormData(); formData.append('ajax_action', 'delete_template'); formData.append('id', id);
            fetch('', { method: 'POST', body: formData }).then(r => r.json()).then(res => { if(res.status === 'success') loadSafetyTemplates(); });
        }

        function renderTemplateTable() {
            const tbody = document.getElementById('templateListTable'); tbody.innerHTML = '<tr><th style="width: auto;">タイトル</th><th style="width: 100px;">操作</th></tr>';
            safetyTemplates.forEach(t => { tbody.innerHTML += `<tr><td style="text-align:left; font-weight:bold; white-space:normal;">${t.title}</td><td><button class="btn btn-blue" style="padding:2px 5px;" onclick="editTemplate(${t.id})">編集</button> <button class="btn btn-red" style="padding:2px 5px;" onclick="deleteTemplate(${t.id})">削除</button></td></tr>`; });
        }

        function updateSaveName() {
            const job = document.getElementById('input_job_no').value || '未定'; document.getElementById('plan_save_name').value = job;
        }

        function syncNightLeader(day) {
            const leaderSel = document.getElementsByName('our_leader_' + day)[0]; const leaderVal = leaderSel.value; if(!leaderVal) return;
            const st = document.getElementsByName('start_' + day)[0].value; let isNight = false;
            if(st) { let hour = parseInt(st.split(':')[0], 10); if(hour <= 5 || hour >= 18) isNight = true; }
            if(isNight) { const clSel = document.getElementsByName('our_cl_' + day)[0]; clSel.value = leaderVal; clSel.dispatchEvent(new Event('change')); }
        }

        function clearDay(btn, day) {
            if(!confirm(`${day}日目のこの項目の入力内容をクリアしますか？`)) return;
            let table = btn.closest('table');
            table.querySelectorAll(`[name$="_${day}"]`).forEach(el => {
                if (el.type === 'checkbox') { el.checked = false; } else { el.value = ''; }
                if(el.tagName === 'SELECT') el.dispatchEvent(new Event('change'));
            });
        }

        function updateWorkerButtons() {
            document.getElementById('addWorkerBtn').style.display = (workerCols < 4) ? 'inline-block' : 'none';
            document.getElementById('removeWorkerBtn').style.display = (workerCols > 1) ? 'inline-block' : 'none';
        }

        function addWorkerCol() {
            if(workerCols < 4) { workerCols++; document.querySelectorAll('.cw' + workerCols).forEach(el => el.style.display = 'table-cell'); updateWorkerButtons(); }
        }

        function removeWorkerCol() {
            if(workerCols > 1) {
                for(let day=1; day<=5; day++) { document.getElementsByName(`our_w${workerCols}_${day}`)[0].value = ""; }
                document.querySelectorAll('.cw' + workerCols).forEach(el => el.style.display = 'none'); workerCols--; updateWorkerButtons();
            }
        }

        function setTime(btn, start, end, day) {
            const td = btn.closest('td'); td.querySelector('input[name^="start_"]').value = start; td.querySelector('input[name^="end_"]').value = end; if(day) syncNightLeader(day);
        }

        function dragStart(e) { e.dataTransfer.setData("text/plain", e.target.closest('td').dataset.day); e.dataTransfer.effectAllowed = "copy"; }
        function allowDrop(e) { e.preventDefault(); e.target.closest('td').classList.add('drag-over'); }
        function dragLeave(e) { e.target.closest('td').classList.remove('drag-over'); }
        
        function dropCopy(e) {
            e.preventDefault(); let td = e.target.closest('td'); td.classList.remove('drag-over');
            let srcDay = e.dataTransfer.getData("text/plain"); let tgtDay = td.dataset.day;
            if (srcDay && tgtDay && srcDay !== tgtDay) {
                let table = td.closest('table'); let inputs = table.querySelectorAll(`[name$="_${srcDay}"]`);
                inputs.forEach(srcInput => {
                    let tgtName = srcInput.name.replace(`_${srcDay}`, `_${tgtDay}`); let tgtInput = table.querySelector(`[name="${tgtName}"]`);
                    if (tgtInput) {
                        if (tgtInput.type === 'checkbox') { tgtInput.checked = srcInput.checked; } else { tgtInput.value = srcInput.value; }
                        if(tgtInput.tagName === 'SELECT') tgtInput.dispatchEvent(new Event('change'));
                    }
                });
            }
        }

        function openDbModal() { document.getElementById('dbModal').style.display='block'; renderPersonTable(); }

        function savePerson() {
            const formData = new FormData(); formData.append('ajax_action', 'save_person'); formData.append('type', document.getElementById('db_type').value);
            formData.append('company', document.getElementById('db_company').value); formData.append('name', document.getElementById('db_name').value); formData.append('phone', document.getElementById('db_phone').value);
            fetch('', { method: 'POST', body: formData }).then(r => r.json()).then(res => {
                if(res.status === 'success') { document.getElementById('db_company').value = ''; document.getElementById('db_name').value = ''; document.getElementById('db_phone').value = ''; loadMasterData(); }
            });
        }

        function importCSV() {
            const fileInput = document.getElementById('csv_file'); if(!fileInput.files.length) { alert('CSVファイルを選択してください。'); return; }
            const file = fileInput.files[0]; const reader = new FileReader();
            reader.onload = function(e) {
                const text = e.target.result; const rows = text.split('\n').map(row => row.split(','));
                const formData = new FormData(); formData.append('ajax_action', 'import_csv'); formData.append('csv_data', JSON.stringify(rows));
                fetch('', { method: 'POST', body: formData }).then(r => r.json()).then(res => {
                    if(res.status === 'success') { alert('CSVから名簿を一括登録しました！'); fileInput.value = ''; loadMasterData(); }
                });
            };
            reader.readAsText(file, 'Shift_JIS');
        }

        function deletePerson(id) {
            if(!confirm('この名簿データを削除しますか？')) return;
            const formData = new FormData(); formData.append('ajax_action', 'delete_person'); formData.append('id', id);
            fetch('', { method: 'POST', body: formData }).then(r => r.json()).then(res => { if(res.status === 'success') loadMasterData(); });
        }

        function loadMasterData() {
            const formData = new FormData(); formData.append('ajax_action', 'get_personnel');
            fetch('', { method: 'POST', body: formData }).then(r => r.json()).then(data => { masterData = data; buildSelects(); renderPersonTable(); });
        }

        function renderPersonTable() {
            const tbody = document.getElementById('personListTable'); tbody.innerHTML = '<tr><th style="width: 60px;">区分</th><th style="width: auto;">業者名</th><th style="width: 100px;">氏名</th><th style="width: 110px;">携帯番号</th><th style="width: 60px;">操作</th></tr>';
            masterData.forEach(p => {
                let typeStr = p.type === 'our' ? '<span style="color:blue;">当社</span>' : '<span style="color:green;">協力</span>';
                tbody.innerHTML += `<tr><td>${typeStr}</td><td style="white-space:normal;">${p.company}</td><td>${p.name}</td><td>${p.phone}</td><td><button class="btn btn-red" style="padding:2px 5px;" onclick="deletePerson(${p.id})">削除</button></td></tr>`;
            });
        }

        function buildSelects() {
            const ourSelects = document.querySelectorAll('.sel_our'); let ourHtml = '<option value="">選択</option>';
            masterData.filter(p => p.type === 'our').forEach(p => { ourHtml += `<option value="${p.name}" data-phone="${p.phone}">${p.name}</option>`; });
            ourSelects.forEach(sel => { let val = sel.value; sel.innerHTML = ourHtml; sel.value = val; });

            const compSelects = document.querySelectorAll('.sel_company');
            let companies = [...new Set(masterData.filter(p => p.type === 'partner').map(p => p.company))]; let compHtml = '<option value="">選択</option>';
            companies.forEach(c => { if(c) compHtml += `<option value="${c}">${c}</option>`; });
            
            compSelects.forEach(sel => { let val = sel.value; sel.innerHTML = compHtml; sel.value = val; if(val) sel.dispatchEvent(new Event('change')); });
        }

        function filterPartner(day) {
            const compName = document.getElementsByName('part_name_'+day)[0].value; const leaderSel = document.getElementById('pl_'+day); const phoneInput = document.getElementById('part_phone_'+day);
            let currentLeader = leaderSel.value; leaderSel.innerHTML = '<option value="">選択</option>'; phoneInput.value = '';
            if(compName) { masterData.filter(p => p.type === 'partner' && p.company === compName).forEach(p => { leaderSel.innerHTML += `<option value="${p.name}" data-phone="${p.phone}">${p.name}</option>`; }); leaderSel.value = currentLeader; }
        }

        function setPhone(selObj, phoneInputId) {
            const phoneInput = document.getElementById(phoneInputId); const opt = selObj.options[selObj.selectedIndex];
            if(opt && opt.dataset.phone) { phoneInput.value = opt.dataset.phone; } else { phoneInput.value = ''; }
        }

        // ==========================================
        // ★ 保存・読込関連のコア関数
        // ==========================================
        
        function handleManualSave() {
            savePlanToDB(false, () => {
                alert('データベースに保存しました。');
            });
        }

        function savePlanToDB(isOverwrite = false, callback = null) {
            const title = document.getElementById('plan_save_name').value;
            if(!title) { alert('保存名を入力してください。'); return; }

            const formElem = document.getElementById('planForm');
            const formDataObj = new FormData(formElem);
            
            let jsonData = {};
            for(let [key, val] of formDataObj.entries()) {
                if(key === 'generate_excel') continue;
                if(jsonData[key]) {
                    if(!Array.isArray(jsonData[key])) jsonData[key] = [jsonData[key]];
                    jsonData[key].push(val);
                } else {
                    jsonData[key] = val;
                }
            }

            const req = new FormData();
            if (isOverwrite && loadedPlanId) {
                req.append('ajax_action', 'overwrite_plan');
                req.append('id', loadedPlanId);
            } else {
                req.append('ajax_action', 'save_plan');
            }
            req.append('title', title);
            req.append('form_data', JSON.stringify(jsonData));

            fetch('', { method: 'POST', body: req })
                .then(r => {
                    if(!r.ok) throw new Error('通信エラー');
                    return r.json();
                })
                .then(res => {
                    if(res.status === 'success') {
                        document.getElementById('plan_save_name').value = res.saved_title;
                        loadedPlanId = res.id;
                        loadedTitle = res.saved_title;
                        for(let i=1; i<=5; i++) loadedDates[i-1] = document.getElementsByName('date_'+i)[0].value;
                        loadPlanList();
                        if (callback) callback();
                    } else {
                        alert('保存エラー: ' + res.message);
                        if (callback) callback(); 
                    }
                })
                .catch(err => {
                    console.error(err);
                    alert('⚠️ DB保存中にエラーが発生しました。\n保存はスキップしてExcelのダウンロードのみ実行します。');
                    if (callback) callback(); 
                });
        }

        function loadPlanList() {
            const req = new FormData();
            req.append('ajax_action', 'get_plans');
            fetch('', { method: 'POST', body: req })
                .then(r => r.json())
                .then(data => {
                    const sel = document.getElementById('load_plan_select');
                    const currentVal = sel.value;
                    sel.innerHTML = '<option value="">過去のデータを読み込む...</option>';
                    data.forEach(p => {
                        const opt = document.createElement('option');
                        opt.value = p.id;
                        opt.text = p.created_at.split(' ')[0] + ' : ' + p.title;
                        sel.appendChild(opt);
                    });
                    if(currentVal) sel.value = currentVal;
                });
        }

        function loadPlanFromDB() {
            const id = document.getElementById('load_plan_select').value;
            if(!id) return;

            const req = new FormData();
            req.append('ajax_action', 'load_plan');
            req.append('id', id);
            fetch('', { method: 'POST', body: req })
                .then(r => r.json())
                .then(data => {
                    if(data.form_data) {
                        const parsed = JSON.parse(data.form_data);
                        document.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
                        document.getElementById('danger_other_text').style.display = 'none';
                        document.getElementById('csv_restore_msg').style.display = 'none';

                        for(let key in parsed) {
                            if(key === 'dangers[]') {
                                let values = Array.isArray(parsed[key]) ? parsed[key] : [parsed[key]];
                                values.forEach(v => {
                                    const cb = document.querySelector(`input[name="dangers[]"][value="${v}"]`);
                                    if(cb) {
                                        cb.checked = true;
                                        if(v === 'その他') document.getElementById('danger_other_text').style.display = 'inline-block';
                                    }
                                });
                                continue;
                            }
                            
                            const el = document.getElementsByName(key)[0];
                            if(el) {
                                if (el.type === 'checkbox') {
                                    el.checked = (parsed[key] === el.value);
                                } else {
                                    el.value = parsed[key];
                                    if(el.tagName === 'SELECT') {
                                        el.dispatchEvent(new Event('change'));
                                        if(key.startsWith('part_leader')) {
                                            setTimeout(() => { el.value = parsed[key]; el.dispatchEvent(new Event('change')); }, 100);
                                        }
                                    }
                                    if(el.tagName === 'TEXTAREA') autoResize(el);
                                }
                            }
                        }
                        
                        loadedPlanId = id;
                        for(let i=1; i<=5; i++) {
                            loadedDates[i-1] = document.getElementsByName('date_'+i)[0].value;
                        }
                        
                        const selText = document.getElementById('load_plan_select').options[document.getElementById('load_plan_select').selectedIndex].text;
                        const titlePart = selText.split(' : ')[1];
                        if(titlePart) {
                            document.getElementById('plan_save_name').value = titlePart;
                            loadedTitle = titlePart;
                        }
                        
                        if(parsed.yorudatsu_csv_data && parsed.yorudatsu_csv_data.length > 5) {
                            document.getElementById('csv_restore_msg').style.display = 'inline-block';
                            document.getElementById('csv_input_file').value = ''; 
                        }
                        
                        workerCols = 1;
                        for(let i=2; i<=4; i++) {
                            let hasValue = false;
                            for(let day=1; day<=5; day++) {
                                if(document.getElementsByName(`our_w${i}_${day}`)[0].value) hasValue = true;
                            }
                            if(hasValue) workerCols = i;
                        }
                        document.querySelectorAll('.cw2, .cw3, .cw4').forEach(e => e.style.display = 'none');
                        for(let i=2; i<=workerCols; i++) {
                            document.querySelectorAll('.cw' + i).forEach(e => e.style.display = 'table-cell');
                        }
                        updateWorkerButtons();
                        alert('データを復元しました。');
                    }
                });
        }

        function deletePlanFromDB() {
            const id = document.getElementById('load_plan_select').value;
            if(!id) { alert('削除するデータを選択してください。'); return; }
            if(!confirm('選択した保存データを削除しますか？\n（この操作は取り消せません）')) return;

            const req = new FormData();
            req.append('ajax_action', 'delete_plan');
            req.append('id', id);
            fetch('', { method: 'POST', body: req })
                .then(r => r.json())
                .then(res => {
                    if(res.status === 'success') {
                        alert('データを削除しました。');
                        document.getElementById('plan_save_name').value = '';
                        loadedPlanId = null;
                        loadedTitle = '';
                        loadPlanList();
                    }
                });
        }
    </script>
</body>
</html>
