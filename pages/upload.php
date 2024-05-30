<?php
require_once '../vendor/autoload.php'; // Include PhpSpreadsheet library

// Database connection
$host = 'localhost';
$db = 'dashboard';
$user = 'root';
$pass = '';
$charset = 'utf8mb4';

$dsn = "mysql:host=$host;dbname=$db;charset=$charset";
$options = [
    PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
    PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
    PDO::ATTR_EMULATE_PREPARES   => false,
];

try {
    $pdo = new PDO($dsn, $user, $pass, $options);
} catch (\PDOException $e) {
    echo json_encode(['error' => 'Database connection failed: ' . $e->getMessage()]);
    exit();
}

$target_dir = "upload/";

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (isset($_FILES['fileInput']) && $_FILES['fileInput']['error'] === UPLOAD_ERR_OK) {
        $filename = $_FILES['fileInput']['name'];
        $tmp_name = $_FILES['fileInput']['tmp_name'];
        $extension = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        $target_file = $target_dir . basename($filename);

        // Check if file already exists
        if (file_exists($target_file)) {
            echo json_encode(['error' => 'File already exists.']);
            exit();
        }

        // Check file size
        if ($_FILES['fileInput']['size'] > 50000000) {
            echo json_encode(['error' => 'File is too large.']);
            exit();
        }

        // Allow only XLS and XLSX files
        if (!in_array($extension, ['xls', 'xlsx'])) {
            echo json_encode(['error' => 'Only XLS and XLSX files are allowed.']);
            exit();
        }

        // Upload file to server
        if (!move_uploaded_file($tmp_name, $target_file)) {
            echo json_encode(['error' => 'Error uploading file.']);
            exit();
        }

        // Process the uploaded spreadsheet
        try {
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($target_file);
            $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);

            // Store file information in the database
            $stmt = $pdo->prepare('INSERT INTO uploads (file_name, file_path) VALUES (?, ?)');
            $stmt->execute([$filename, $target_file]);

            echo json_encode(['success' => 'File uploaded and processed successfully.', 'data' => $sheetData]);
        } catch (Exception $e) {
            echo json_encode(['error' => 'Error processing spreadsheet: ' . $e->getMessage()]);
        }
    } else {
        // Handle file upload errors
        $uploadErrorMessages = [
            UPLOAD_ERR_INI_SIZE   => 'The uploaded file exceeds the upload_max_filesize directive in php.ini.',
            UPLOAD_ERR_FORM_SIZE  => 'The uploaded file exceeds the MAX_FILE_SIZE directive specified in the HTML form.',
            UPLOAD_ERR_PARTIAL    => 'The uploaded file was only partially uploaded.',
            UPLOAD_ERR_NO_FILE    => 'No file was uploaded.',
            UPLOAD_ERR_NO_TMP_DIR => 'Missing a temporary folder.',
            UPLOAD_ERR_CANT_WRITE => 'Failed to write file to disk.',
            UPLOAD_ERR_EXTENSION  => 'A PHP extension stopped the file upload.'
        ];
        $error_code = $_FILES['fileInput']['error'];
        $error_message = $uploadErrorMessages[$error_code] ?? 'Unknown upload error.';
        echo json_encode(['error' => $error_message]);
    }
} else {
    echo json_encode(['error' => 'Invalid request method.']);
}

$pdo = null;
?>
