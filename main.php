<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Получение основных данных
    $lastname = $_POST['lastname'] ?? '';
    $city = $_POST['city'] ?? '';
    $deliveryDate = $_POST['delivery-date'] ?? '';
    $address = $_POST['address'] ?? '';
    $color = $_POST['color'] ?? '';
    // Получение выбранной мебели
    // checkbox имеет имя начинающееся с furniture_. Необходимо перебрать все такие элементы и поместить их в массив.
    $selectedFurniture = [];
    foreach ($_POST as $key => $value) {
        if (str_starts_with($key, 'furniture_')) {
            $selectedFurniture[] = $value;
        }
    }
    // Получение количества для каждого предмета
    $quantities = [];
    for ($i = 1; $i <= 6; $i++) {
        $quantities[$i] = $_POST["quantity_$i"] ?? '0';
    }

    $uploadFile = '';
    // Проверяем, был ли файл загружен без ошибок
    if (isset($_FILES['price_file']) && $_FILES['price_file']['error'] === UPLOAD_ERR_OK) {
        // Папка, куда будут сохраняться загруженные файлы
        $uploadDir = __DIR__ . '/uploads/';

        // Создаем папку, если её нет
        if (!file_exists($uploadDir)) {
            mkdir($uploadDir, 0755, true);
        }
        $uploadFile = $uploadDir . basename($_FILES['price_file']['name']);
        // Перемещаем файл из временной директории в постоянное место хранения
        if (move_uploaded_file($_FILES['price_file']['tmp_name'], $uploadFile)) {
            # echo "Файл успешно загружен.";
        } else {
            echo "Ошибка при сохранении файла.";
        }
    } else {
        echo "Ошибка загрузки файла.";
    }

    $prices = [];

    $phpWord = PhpOffice\PhpWord\IOFactory::load($uploadFile);
    $sections = $phpWord->getSections();

    foreach ($sections as $section) {
        $elements = $section->getElements();

        // Check each element
        foreach ($elements as $element) {
            if ($element instanceof \PhpOffice\PhpWord\Element\Table) {
                $rows = $element->getRows();
                for ($i = 2; $i < count($rows); $i++) {
                    $prices[$i - 2][0] = $rows[$i]->getCells()[0]->getElements()[0]->getText();
                    $prices[$i - 2][1] = (int) $rows[$i]->getCells()[1]->getElements()[0]->getText();
                }
            }
        }
    }

    $number = random_int(1000, 9999);

    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->mergeCells('B2:G2');
    $sheet->mergeCells('B3:G3');
    $sheet->mergeCells('B4:G4');
    $sheet->mergeCells('B5:G5');

    $sheet->getStyle('B2')->getAlignment()->setHorizontal('center');
    $sheet->getStyle('B2')->getFont()->setBold(true);
    $sheet->getStyle('B3')->getAlignment()->setHorizontal('center');
    $sheet->getStyle('B4')->getAlignment()->setHorizontal('center');
    $sheet->getStyle('B5')->getAlignment()->setHorizontal('center');
    $sheet->getStyle('B7')->getAlignment()->setHorizontal('center');
    $sheet->getStyle('B7:G7')->getFont()->setBold(true);

    $sheet->getRowDimension(1)->setRowHeight(30);

    $drawing = new Drawing();
    $drawing->setPath('./assets/штрих.JPG');
    $drawing->setHeight(40);
    $drawing->setCoordinates('F1');
    $drawing->setWorksheet($spreadsheet->getActiveSheet());

    $sheet->getColumnDimension('C')->setWidth(30);
    $sheet->getColumnDimension('F')->setWidth(20);

    $data = [
        ['Накладная №' . $number],
        ['Адрес доставки: ' . $city . ', ' . $address],
        ['Дата доставки: ' . $deliveryDate],
        ['Получатель: ' . $lastname],
        [],
        ['№', 'Наименование товара', 'Цвет', 'Цена', 'Количество', 'Сумма']
    ];

    $orderedItems = [];

    foreach ($selectedFurniture as $index => $furniture) {
        switch ($furniture) {
            case 'Банкетка':
                $orderedItems[] = ["Банкетка", $prices[0][1], $quantities[$index + 1]];
                break;
            case 'Кровать':
                $orderedItems[] = ["Кровать", $prices[1][1], $quantities[$index + 1]];
                break;
            case 'Комод':
                $orderedItems[] = ["Комод", $prices[2][1], $quantities[$index + 1]];
                break;
            case 'Шкаф':
                $orderedItems[] = ["Шкаф", $prices[3][1], $quantities[$index + 1]];
                break;
            case 'Стул':
                $orderedItems[] = ["Стул", $prices[4][1], $quantities[$index + 1]];
                break;
            case 'Стол':
                $orderedItems[] = ["Стол", $prices[5][1], $quantities[$index + 1]];
                break;
        }
    }

    $sum = 0;
    for ($i = 1; $i < count($orderedItems); $i++) {
        $subsum = $orderedItems[$i][1] * $orderedItems[$i][2];
        $sum += $subsum;

        $data[] = [$i, $orderedItems[$i][0], '', $orderedItems[$i][1], $orderedItems[$i][2], $subsum];
    }

    $mult = 1;
    switch ($color) {
        case 'Орех':
            $mult = 1.1;
            break;
        case 'Дуб морёный':
            $mult = 1.2;
            break;
        case 'Палисандр':
            $mult = 1.3;
            break;
        case 'Эбеновое дерево':
            $mult = 1.4;
            break;
        case 'Клён':
            $mult = 1.5;
            break;
        case 'Лиственница':
            $mult = 1.6;
            break;
    }

    $nMerge = 7 + $i;
    $sheet->mergeCells('E' . $nMerge . ':F' . $nMerge, 'B' . $nMerge + 1 . ':F' . $nMerge + 1);
    $sheet->getRowDimension($nMerge)->setRowHeight(40);


    $drawing = new Drawing();
    $drawing->setPath('./assets/' . mb_strtolower($color) . '.png');
    $drawing->setHeight(50);
    $drawing->setCoordinates('E' . $nMerge);
    $drawing->setWorksheet($spreadsheet->getActiveSheet());

    $data[] = ['', 'Цвет: ' . $color, '', '', '', $sum];
    $data[] = ['Итого', '', $sum * $mult];

    // Write data to cells
    $row = 2;
    foreach ($data as $rowData) {
        $col = 'B';
        foreach ($rowData as $cellData) {
            $sheet->setCellValue($col . $row, $cellData);
            $col++;
        }
        $row++;
    }

    $writer = new Xlsx($spreadsheet);
    $filename = 'Документ_на_выдачу_№' . $number . '.xlsx';

    $writer->save($filename);

    $output = "<a href='./$filename'>$filename</a>";
    echo $output;
}

?>