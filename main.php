<?php

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Получение основных данных
    $lastname = $_POST['lastname'] ?? '';
    $city = $_POST['city'] ?? '';
    $deliveryDate = $_POST['delivery-date'] ?? '';
    $address = $_POST['address'] ?? '';
    $color = $_POST['color'] ?? '';
    // Получение выбранной мебели
    $selectedFurniture = $_POST['furniture'] ?? [];
    // Получение количества для каждого предмета
    $quantities = [];
    for ($i = 1; $i <= 6; $i++) {
        $quantities[$i] = $_POST["quantity_$i"] ?? '0';
    }

    // Проверяем, был ли файл загружен без ошибок
    if (isset($_FILES['priceFile']) && $_FILES['priceFile']['error'] === UPLOAD_ERR_OK) {
        // Папка, куда будут сохраняться загруженные файлы
        $uploadDir = __DIR__ . '/uploads/';

        // Создаем папку, если её нет
        if (!file_exists($uploadDir)) {
            mkdir($uploadDir, 0755, true);
        }
        $uploadFile = $uploadDir . basename($_FILES['priceFile']['name']);
        // Перемещаем файл из временной директории в постоянное место хранения
        if (move_uploaded_file($_FILES['priceFile']['tmp_name'], $uploadFile)) {
            echo "Файл успешно загружен.";
        } else {
            echo "Ошибка при сохранении файла.";
        }
    } else {
        echo "Ошибка загрузки файла.";
    }


}

?>