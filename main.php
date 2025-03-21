<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Border;

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
        if (isset($_POST["quantity_$i"]) && $_POST["quantity_$i"] !== '')
            $quantities[] = $_POST["quantity_$i"];
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
                            $prices[$i - 2][1] = (int)$rows[$i]->getCells()[1]->getElements()[0]->getText();
                        }
                    }
                }
            }

            $number = random_int(1000, 9999);

            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $spreadsheet->getDefaultStyle()->getFont()->setName('Times New Roman');

            $sheet->mergeCells('B2:G2');
            $sheet->mergeCells('B3:G3');
            $sheet->mergeCells('B4:G4');
            $sheet->mergeCells('B5:G5');

            $sheet->getStyle('B2')->getFont()->setBold(true);
            $sheet->getStyle('B2:B6')->getAlignment()->setHorizontal('center');
            $sheet->getStyle('B7:G7')->getFont()->setBold(true);
            $sheet->getStyle('B7:G7')->getAlignment()->setHorizontal('center');

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
                        $orderedItems[] = ["Банкетка", $prices[0][1], $quantities[$index]];
                        break;
                    case 'Кровать':
                        $orderedItems[] = ["Кровать", $prices[1][1], $quantities[$index]];
                        break;
                    case 'Комод':
                        $orderedItems[] = ["Комод", $prices[2][1], $quantities[$index]];
                        break;
                    case 'Шкаф':
                        $orderedItems[] = ["Шкаф", $prices[3][1], $quantities[$index]];
                        break;
                    case 'Стул':
                        $orderedItems[] = ["Стул", $prices[4][1], $quantities[$index]];
                        break;
                    case 'Стол':
                        $orderedItems[] = ["Стол", $prices[5][1], $quantities[$index]];
                        break;
                }
            }

            $sum = 0;
            for ($i = 0; $i < count($orderedItems); $i++) {
                $subsum = $orderedItems[$i][1] * $orderedItems[$i][2];
                $sum += $subsum;
                $data[] = [$i + 1, $orderedItems[$i][0], '', $orderedItems[$i][1], $orderedItems[$i][2], $subsum];
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

            $nMerge = 7 + $i + 1;
            $sheet->mergeCells('E' . $nMerge . ':F' . $nMerge);
            $sheet->mergeCells('B' . $nMerge + 1 . ':F' . $nMerge + 1);
            $sheet->getStyle('B8:' . 'B' . $nMerge + 1)->getAlignment()->setHorizontal('center');
            $sheet->getStyle('E8:' . 'G' . $nMerge + 1)->getAlignment()->setHorizontal('center');
            $sheet->getStyle('D' . $nMerge . ':F' . $nMerge)->getAlignment()->setHorizontal('center');
            $sheet->getStyle('B' . $nMerge . ':F' . $nMerge)->getAlignment()->setVertical('center');
            $sheet->getStyle('B' . $nMerge + 1 . ':F' . $nMerge + 1)->getFont()->setBold(true);
            $sheet->getRowDimension($nMerge)->setRowHeight(40);

            $drawing = new Drawing();
            $drawing->setPath('./assets/' . mb_strtolower($color) . '.png');
            $drawing->setHeight(45);
            $drawing->setCoordinates('D' . $nMerge);
            $drawing->setOffsetX(10);
            $drawing->setOffsetY(5);
            $drawing->setWorksheet($spreadsheet->getActiveSheet());

            $data[] = ['', 'Цвет: ' . $color, '', $mult, '', $sum];
            $data[] = ['Итого', '', '', '', '', $sum * $mult];

            $totalItems = count($orderedItems);
            $finalSum = $sum * $mult;
            $sheet->setCellValue('B' . ($nMerge + 3), "Всего наименований {$totalItems}, на сумму {$finalSum} руб.");

            $styleArray = [
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THICK, // Толстая граница
                    ],
                ],
            ];

            $sheet->getStyle('B7:G' . $nMerge)->applyFromArray($styleArray);
            $sheet->getStyle('G' . $nMerge + 1)->applyFromArray($styleArray);

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

            $warrantyText = file_get_contents('./assets/Гарантийное обслуживание.txt');
            $warrantyLines = explode("\n", $warrantyText);

            $currentRow = $nMerge + 5;
            foreach ($warrantyLines as $index => $line) {
                $sheet->mergeCells("B{$currentRow}:G{$currentRow}");
                $sheet->setCellValue("B{$currentRow}", trim($line));
                $sheet->getRowDimension($currentRow)->setRowHeight(30);
                $sheet->getStyle("B{$currentRow}")->getAlignment()->setWrapText(true);
                if ($index === 0) {
                    $sheet->getStyle("B{$currentRow}")->getFont()->setBold(true);
                }
                $currentRow++;
            }

            $writer = new Xlsx($spreadsheet);
            $filename = 'Документ_на_выдачу_№' . $number . '.xlsx';

            $writer->save($filename);

            echo "<div class='success-message'>Заказ №$number успешно оформлен.<br>
                  Общая сумма заказа: $finalSum руб.<br>
                  Количество наименований: $totalItems<br>
                  Скачать документ <a href='./$filename'>$filename</a></div>";
        } else {
            echo "<div class='error-message'>Ошибка при сохранении файла</div>";
        }
    } else {
        echo "<div class='error-message'>Ошибка при загрузке файла</div>";
    }
}
?>