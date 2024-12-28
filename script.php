<?php

require($_SERVER['DOCUMENT_ROOT'] . '/bitrix/modules/main/include/prolog_before.php');

if (!$_REQUEST['download']) {
    ?>
    <p style="margin:50px;line-height:24px;text-align:center;">
        <a href="get_employees_empty_departments.php?download=yes">Выгрузить пользователей без подразделения</a>
    </p>
    <?php
    die();
}

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use Bitrix\Main\UserTable;
use Bitrix\Main\UserGroupTable;
use Bitrix\Main\Loader;

Loader::includeModule('main');

$userGroupIds = [];
$userGroups = UserGroupTable::getList([
    'select' => ['USER_ID'],
    'filter' => ['GROUP_ID' => 12],
]);

while ($userGroup = $userGroups->fetch()) {
    $userGroupIds[] = $userGroup['USER_ID'];
}

$existDepartmentIds = [];
$departments = \Bitrix\Iblock\SectionTable::getList([
    'select' => ['ID'],
    'filter' => ['IBLOCK_ID' => 5, 'ACTIVE' => 'Y'],
]);

while ($department = $departments->fetch()) {
    $existDepartmentIds[] = $department['ID'];
}

$users = UserTable::getList([
    'select' => ['ID', 'NAME', 'LAST_NAME', 'SECOND_NAME', 'WORK_POSITION', 'UF_DEPARTMENT', 'XML_ID'],
    'filter' => [
        'ACTIVE' => 'Y',
        '!XML_ID' => '',
        'ID' => $userGroupIds,
        [
            'LOGIC' => 'OR',
            '!UF_DEPARTMENT' => $existDepartmentIds,
            'UF_DEPARTMENT' => false,
        ],
    ],
]);

$rows = [];
$rows[] = ['ID пользователя', 'ФИО пользователя', 'Должность', 'Внешний код'];

while ($user = $users->fetch()) {
    $fullName = trim($user['LAST_NAME'] . ' ' . $user['NAME'] . ' ' . $user['SECOND_NAME']);

    $rows[] = [
        $user['ID'],
        $fullName,
        $user['WORK_POSITION'],
        $user['XML_ID']
    ];
}

// Создание Excel файла
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->fromArray($rows, null, 'A1');

$sheet->getStyle('1:1')->getFont()->setBold(true);
$sheet->getStyle('A:Z')->getAlignment()->setVertical(Alignment::VERTICAL_TOP);

foreach (range('A', 'E') as $columnID) {
    $sheet->getColumnDimension($columnID)->setAutoSize(true);
}

// Сохранение файла
$writer = new Xlsx($spreadsheet);
$filename = 'users_without_department_' . date('Y-m-d') . '.xlsx';

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . $filename . '"');
$writer->save('php://output');
