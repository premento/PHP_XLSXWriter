<?php
// Load the library through composer when available, otherwise directly.
$autoload = __DIR__ . '/../vendor/autoload.php';
if (file_exists($autoload)) {
    require_once $autoload;
} else {
    require_once __DIR__ . '/../xlsxwriter.class.php';
}

// Any notice, warning or deprecation raised by the library fails the test that triggered it.
set_error_handler(function ($severity, $message, $file, $line) {
    if (!(error_reporting() & $severity)) {
        return false;
    }
    throw new ErrorException($message, 0, $severity, $file, $line);
});
error_reporting(E_ALL);
