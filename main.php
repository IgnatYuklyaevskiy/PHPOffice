<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Border;
/**
 * Функция для извлечения цен из загруженного Word-файла.
 */
function extractPricesFromWord($uploadFile) {
    $prices = [];
    $phpWord = IOFactory::load($uploadFile);
    foreach ($phpWord->getSections() as $section) {
        foreach ($section->getElements() as $element) {
            if ($element instanceof \PhpOffice\PhpWord\Element\Table) {
                $rows = $element->getRows();
                // Пропускаем первые две строки (заголовок и, возможно, подзаголовок)
                for ($i = 2; $i < count($rows); $i++) {
                    $cellElements = $rows[$i]->getCells();
                    if (isset($cellElements[0], $cellElements[1])) {
                        $name = $cellElements[0]->getElements()[0]->getText();
                        $price = (int) $cellElements[1]->getElements()[0]->getText();
                        $prices[] = ['name' => $name, 'price' => $price];
                    }
                }
            }
        }
    }
    return $prices;
}
/**
 * Функция для формирования массива заказа.
 */
function buildOrderItems($selectedFurniture, $quantities, $prices) {
    // Преобразуем цены в ассоциативный массив по имени товара
    $priceMapping = [];
    foreach ($prices as $item) {
        $priceMapping[$item['name']] = $item['price'];
    }

    $orderedItems = [];
    foreach ($selectedFurniture as $index => $furniture) {
        if (isset($priceMapping[$furniture]) && isset($quantities[$index])) {
            $orderedItems[] = [
                'name' => $furniture,
                'price' => $priceMapping[$furniture],
                'quantity' => $quantities[$index]
            ];
        }
    }
    return $orderedItems;
}
/**
 * Функция для установки базовых стилей и заголовка документа.
 */
function setupSpreadsheet($spreadsheet, $sheet, $number, $city, $address, $deliveryDate, $lastname) {
    $spreadsheet->getDefaultStyle()->getFont()->setName('Times New Roman');

    // Объединение и заполнение ячеек заголовка
    $sheet->mergeCells('B2:G2');
    $sheet->mergeCells('B3:G3');
    $sheet->mergeCells('B4:G4');
    $sheet->mergeCells('B5:G5');
    $sheet->setCellValue('B2', 'Накладная №' . $number);
    $sheet->setCellValue('B3', 'Адрес доставки: ' . $city . ', ' . $address);
    $sheet->setCellValue('B4', 'Дата доставки: ' . $deliveryDate);
    $sheet->setCellValue('B5', 'Получатель: ' . $lastname);

    $sheet->getStyle('B2')->getFont()->setBold(true);
    $sheet->getStyle('B2:B6')->getAlignment()->setHorizontal('center');
    $sheet->getRowDimension(1)->setRowHeight(30);

    // Настройка ширины столбцов
    $sheet->getColumnDimension('C')->setWidth(30);
    $sheet->getColumnDimension('F')->setWidth(20);
}
/**
 * Функция для вставки данных заказа.
 */
function insertOrderData($sheet, $data, $startRow = 2) {
    $row = $startRow;
    foreach ($data as $rowData) {
        $col = 'B';
        foreach ($rowData as $cellData) {
            $sheet->setCellValue($col . $row, $cellData);
            $col++;
        }
        $row++;
    }
}
/**
 * Функция для вставки рисунка с заданными координатами и смещением.
 */
function placeDrawing($sheet, $path, $height, $coordinate, $offsetX = 0, $offsetY = 0) {
    $drawing = new Drawing();
    $drawing->setPath($path);
    $drawing->setHeight($height);
    $drawing->setCoordinates($coordinate);
    if ($offsetX) {
        $drawing->setOffsetX($offsetX);
    }
    if ($offsetY) {
        $drawing->setOffsetY($offsetY);
    }
    $drawing->setWorksheet($sheet);
}
// Обработка POST-запроса
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Получение основных данных
    $lastname     = $_POST['lastname'] ?? '';
    $city         = $_POST['city'] ?? '';
    $deliveryDate = $_POST['delivery-date'] ?? '';
    $address      = $_POST['address'] ?? '';
    $color        = $_POST['color'] ?? '';

    // Получаем выбранную мебель (checkbox'ы с именами, начинающимися с furniture_)
    $selectedFurniture = [];
    foreach ($_POST as $key => $value) {
        if (str_starts_with($key, 'furniture_')) {
            $selectedFurniture[] = $value;
        }
    }

    // Получение количества для каждого элемента
    $quantities = [];
    for ($i = 1; $i <= 6; $i++) {
        if (!empty($_POST["quantity_$i"])) {
            $quantities[] = $_POST["quantity_$i"];
        }
    }

    if (isset($_FILES['price_file']) && $_FILES['price_file']['error'] === UPLOAD_ERR_OK) {
        $uploadDir = __DIR__ . '/uploads/';
        if (!file_exists($uploadDir)) {
            mkdir($uploadDir, 0755, true);
        }
        $uploadFile = $uploadDir . basename($_FILES['price_file']['name']);
        if (move_uploaded_file($_FILES['price_file']['tmp_name'], $uploadFile)) {
            $prices = extractPricesFromWord($uploadFile);
            $orderedItems = buildOrderItems($selectedFurniture, $quantities, $prices);

            // Расчет итоговой суммы заказа
            $sum = 0;
            $data = [
                // Заголовок таблицы
                ['№', 'Наименование товара', 'Цвет', 'Цена', 'Количество', 'Сумма']
            ];
            foreach ($orderedItems as $index => $item) {
                $subSum = $item['price'] * $item['quantity'];
                $sum += $subSum;
                $data[] = [$index + 1, $item['name'], '', $item['price'], $item['quantity'], $subSum];
            }

            // Множитель в зависимости от цвета
            $multipliers = [
                'Орех'           => 1.1,
                'Дуб морёный'    => 1.2,
                'Палисандр'      => 1.3,
                'Эбеновое дерево'=> 1.4,
                'Клён'           => 1.5,
                'Лиственница'    => 1.6
            ];
            $mult = $multipliers[$color] ?? 1;
            $finalSum = $sum * $mult;

            // Генерация номера накладной
            $number = random_int(1000, 9999);

            // Инициализируем Spreadsheet
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();

            // Установка базовых стилей и заголовка документа
            setupSpreadsheet($spreadsheet, $sheet, $number, $city, $address, $deliveryDate, $lastname);

            // Добавим изображение штрих-кода
            placeDrawing($sheet, './assets/штрих.JPG', 40, 'F1');

            // Запишем данные (таблица с заказом) начиная с 8-й строки
            $startDataRow = 8;
            insertOrderData($sheet, $data, $startDataRow);

            // Применим жирные границы к области заказа
            $styleArray = [
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => Border::BORDER_THICK,
                    ],
                ],
            ];
            $sheet->getStyle("B{$startDataRow}:G" . ($startDataRow + count($data) - 1))
                ->applyFromArray($styleArray);

            // Размещение изображения выбранного цвета
            // Здесь сдвигаем картинку относительно выбранной позиции
            $nMerge = $startDataRow + count($data) + 1;
            placeDrawing($sheet, './assets/' . mb_strtolower($color) . '.png', 45, "D{$nMerge}", 10, 5);

            // Дополнительная информация: цвет, сумма, итого
            $sheet->mergeCells("E{$nMerge}:F{$nMerge}");
            $sheet->setCellValue("E{$nMerge}", "Цвет: {$color}");
            $nMerge++;
            $sheet->mergeCells("B{$nMerge}:G{$nMerge}");
            $sheet->setCellValue("B{$nMerge}", "Итого: " . ($sum * $mult));

            // Вывод итоговой информации ниже таблицы
            $totalItems = count($orderedItems);
            $sheet->setCellValue("B" . ($nMerge + 2), "Всего наименований {$totalItems}, на сумму {$finalSum} руб.");

            // Вставка текста из файла "Гарантийное обслуживание.txt"
            $warrantyText = file_get_contents('./assets/Гарантийное обслуживание.txt');
            $warrantyLines = explode("\n", $warrantyText);
            $currentRow = $nMerge + 4;
            foreach ($warrantyLines as $index => $line) {
                $line = trim($line);
                if (!empty($line)) {
                    $sheet->mergeCells("B{$currentRow}:G{$currentRow}");
                    $sheet->setCellValue("B{$currentRow}", $line);
                    $sheet->getRowDimension($currentRow)->setRowHeight(30);
                    $sheet->getStyle("B{$currentRow}")->getAlignment()->setWrapText(true);
                    // Первая строка делается жирной
                    if ($index === 0) {
                        $sheet->getStyle("B{$currentRow}")->getFont()->setBold(true);
                    }
                    $currentRow++;
                }
            }

            // Сохранение файла
            $writer = new Xlsx($spreadsheet);
            $filename = 'Документ_на_выдачу_№' . $number . '.xlsx';
            $writer->save($filename);
            echo "<div class='success-message'>
                    Заказ №{$number} успешно оформлен.<br>
                    Общая сумма заказа: {$finalSum} руб.<br>
                    Количество наименований: {$totalItems}<br>
                    Скачать документ <a href='./{$filename}'>{$filename}</a>
                  </div>";
        } else {
            echo "<div class='error-message'>Ошибка при сохранении файла</div>";
        }
    } else {
        echo "<div class='error-message'>Ошибка при загрузке файла</div>";
    }
}
?>