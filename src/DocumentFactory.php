<?php
namespace Thunderkiss52\LaravelDocuments;


use Exception;
use Illuminate\Support\Arr;
use Illuminate\Support\Carbon;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\IOFactory;
use Illuminate\Support\Collection;
use Illuminate\Support\Facades\Blade;
use PhpOffice\PhpSpreadsheet\Writer\Pdf;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Html;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use Spatie\TemporaryDirectory\TemporaryDirectory;
use ZipArchive;

final class DocumentFactory
{
    /*public static function getShortName(string $fullName, bool $familystart = false): string
    {
        $words = explode(' ', $fullName);
        if (count($words) > 2) {
            if ($familystart) {
                return $words[0] . ' ' . mb_str_split($words[1])[0] . '.' . mb_str_split($words[2])[0] . '.';
            } else {
                return mb_str_split($words[1])[0] . '.' . mb_str_split($words[2])[0] . '. ' . $words[0];
            }
        } else {
            return $fullName;
        }
    }*/

    // public static function setCheckboxValidation(Worksheet $sheet)
    // {

    // }

    public static function setNumberValidation(Worksheet $sheet, int $min = null, int $max = null): DataValidation
    {
        $validation = new DataValidation;
        $validation->setType(DataValidation::TYPE_DECIMAL); // Тип валидации: числа с плавающей точкой
        $validation->setPromptTitle('Числовое значение'); // Заголовок сообщения
        if ($max || $min) {
            $operator = match (true) {
                ($max && $min) => DataValidation::OPERATOR_BETWEEN,
                $min => DataValidation::OPERATOR_GREATERTHANOREQUAL,
                $max => DataValidation::OPERATOR_LESSTHANOREQUAL,
            };

            $txt = match (true) {
                ($max && $min) => "Допустимы только числа от {$min} до {$max}",
                $min => "Введите число больше чем {$min}",
                $max => "Введите число меньше чем {$max}",
            };
            $validation->setError($txt);

            $validation->setOperator($operator);
            if ($min) {
                $validation->setFormula1($min); // Минимальное значение
            }
            if ($max) {
                $validation->setFormula2($max); // Максимальное значение
            }
            $validation->setPrompt($txt); // Текст сообщения

        } else {
            $validation->setPrompt($txt ?? 'Введите число');
        }
        $validation->setShowInputMessage(true); // Показывать сообщение при выборе ячейки
        $validation->setShowErrorMessage(true);
        $validation->setErrorTitle('Ошибка ввода');

        return $validation;
    }
    public static function setMaxTextLengthValidation(Worksheet $sheet, int $length): DataValidation
    {

        $validation = new DataValidation;
        $validation->setType(DataValidation::TYPE_TEXTLENGTH); // Тип валидации: длина текста
        $validation->setOperator(DataValidation::OPERATOR_LESSTHANOREQUAL); // Оператор: меньше или равно
        $validation->setFormula1($length); // Максимальная длина строки (10 символов)
        $validation->setAllowBlank(true); // Разрешить пустые значения
        $validation->setShowInputMessage(true); // Показывать сообщение при выборе ячейки
        $validation->setPromptTitle('Ограничение длины'); // Заголовок сообщения
        $validation->setPrompt("Введите не более {$length} символов"); // Текст сообщения
        $validation->setShowErrorMessage(true); // Показывать сообщение об ошибке
        $validation->setErrorTitle('Ошибка ввода'); // Заголовок ошибки
        $validation->setError("Длина текста не должна превышать {$length} символов"); // Текст ошибки
        return $validation;
    }
    public static function setTimeValidation(Worksheet $sheet): DataValidation
    {
        // Создаем валидацию для выбора времени
        $validation = new DataValidation;
        //NumberFormat::FORMAT_DATE_TIME4
        $validation->setType(DataValidation::TYPE_TIME); // Тип валидации — пользовательский
        $validation->setErrorStyle(DataValidation::STYLE_STOP); // Стиль ошибки
        $validation->setAllowBlank(false); // Запрет на пустые значения
        $validation->setShowInputMessage(true); // Показывать сообщение при выборе ячейки
        $validation->setShowErrorMessage(true); // Показывать сообщение об ошибке
        $validation->setErrorTitle('Ошибка ввода'); // Заголовок ошибки
        $validation->setError('Введите время в формате ЧЧ:ММ (например, 14:30).'); // Сообщение об ошибке
        $validation->setPromptTitle('Выбор времени'); // Заголовок подсказки
        $validation->setPrompt('Введите время в формате ЧЧ:ММ (например, 09:00).'); // Подсказка

        // Устанавливаем формулу для валидации времени (регулярное выражение)
        $validation->setFormula1('=AND(HOUR(E2)<24, MINUTE(E2)<60)');
        return $validation;
    }
    public static function setDateValidation(Worksheet $sheet, string $title, ?Carbon $before = null, ?Carbon $after = null): DataValidation
    {
        $validation = new DataValidation;
        $validation->setType(DataValidation::TYPE_DATE); // Тип валидации — дата
        $validation->setErrorStyle(DataValidation::STYLE_STOP); // Стиль ошибки
        $validation->setAllowBlank(false); // Запрет на пустые значения
        $validation->setShowInputMessage(true); // Показывать сообщение при выборе ячейки
        $validation->setShowErrorMessage(true); // Показывать сообщение об ошибке
        $validation->setErrorTitle('Ошибка ввода'); // Заголовок ошибки
        $validation->setError('Введите корректную дату.'); // Сообщение об ошибке
        $validation->setPromptTitle('Выбор даты'); // Заголовок подсказки
        $validation->setPrompt('Выберите дату в формате ДД.ММ.ГГГГ.'); // Подсказка

        // Устанавливаем диапазон допустимых дат (необязательно)
        if ($before) {
            $validation->setFormula1('DATE(' . $before->format('Y,m,d') . ')'); // Перед датой
        }

        if ($after) {
            $validation->setFormula2('DATE(' . $after->format('Y,m,d') . ')'); // После даты
        }

        return $validation;
    }

    public static function setRadioValidation(Spreadsheet $spreadsheet, Worksheet $sheet, string $title, array $values): DataValidation|null
    {
        // Создаем выпадающий список для столбца "Category"
        $validation = new DataValidation;
        $validation->setType(DataValidation::TYPE_LIST);
        $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);
        $validation->setAllowBlank(false);
        $validation->setShowInputMessage(true);
        $validation->setShowErrorMessage(true);
        $validation->setShowDropDown(true);


        $validation->setPromptTitle($title); // Заголовок подсказки
        $validation->setPrompt('Выберите одно значение из выпадающего списка'); // Подсказка
        // $validation->setErrorTitle('Input error');
        // $validation->setError('Value is not in list.');
        // $validation->setPrompt('Please pick a value from the drop-down list.');
        $validation->setPromptTitle($title);
        $data = implode(',', array_map(fn($v) => str_replace([''], '', $v), $values));
        if (mb_strlen($data) > 255) {
            return self::setDropDownValidation($spreadsheet, $sheet, $title, $values);
        }
        $validation->setFormula1('"' . $data . '"'); // Список значений
        // throw new Exception($data);
        return $validation;
    }
    public static function setMultipleValidation(Spreadsheet $spreadsheet, $options = null): DataValidation|null
    {
        // Создаем выпадающий список для столбца "Category"
        $validation = new DataValidation;
        $validation->setShowInputMessage(true);
        $validation->setShowErrorMessage(true);
        $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);
        $validation->setPromptTitle("Введите несколько фраз через разделительный знак ;"); // Заголовок подсказки
        if ($options) {
            $validation->setPrompt(collect($options)->pluck('label')->implode('; ')); // Подсказка
        } else {
            $validation->setPrompt("Введите любые данные");
        }
        //  else {
        //     $validation->setPrompt('Введите несколько фраз через разделительный знак ;'); // Подсказка
        // }
        return $validation;
    }
    public static function setDropDownValidation(Spreadsheet $spreadsheet, Worksheet $sheet, string $title, array $values): DataValidation|null
    {
        // Создаем выпадающий список для столбца "Category"
        $validation = new DataValidation;
        $validation->setType(DataValidation::TYPE_LIST);
        $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);
        $validation->setAllowBlank(false);
        $validation->setShowInputMessage(true);
        $validation->setShowErrorMessage(true);
        $validation->setShowDropDown(true);

        $validation->setPromptTitle($title); // Заголовок подсказки
        $validation->setPrompt('Выберите одно значение из выпадающего списка'); // Подсказка
        // $validation->setErrorTitle('Input error');
        // $validation->setError('Value is not in list.');
        // $validation->setPrompt('Please pick a value from the drop-down list.');
        $validation->setPromptTitle($title);
        $data = implode(',', array_map(fn($v) => str_replace([''], '', $v), $values));
        // if (mb_strlen($data) > 255) {
        $dataSheet = $spreadsheet->createSheet();
        $dataSheet->setTitle($title);
        $validation->setFormula1('=' . $title . '!$A:$A');


        for ($i = 0; $i < count($values); $i++) {
            $dataSheet->getCell([1, $i + 1])->setValue($values[$i]);//->setDataValidation(clone $val);
        }
        return $validation;
    }
    /**
     * Summary of table
     * @param array<mixed> $rows
     * @param null|array<mixed> $headers
     * @param string $type
     * @return string
     */
    public static function table(array $rows, array $headers = null, bool $filters = true, array $validation = [], array $width = [], string $type = 'xlsx', bool $borders = true): string
    {
        $spreadsheet = new Spreadsheet();
        $activeWorksheet = $spreadsheet->getActiveSheet();

        //Добавление заголовков
        if ($headers) {
            if (array_is_list($headers)) {
                foreach ($headers as $column => $header) {
                    if (array_key_exists($header, $width)) {
                        $activeWorksheet->getColumnDimensionByColumn($column + 1)->setWidth($width[$header]);
                    }

                    $activeWorksheet->setCellValue([$column + 1, 1], $header);
                    if ($validation && array_key_exists($header, $validation)) {

                        $val = match ($validation[$header]['type']) {
                            "radio" => self::setRadioValidation($spreadsheet, $activeWorksheet, $header, $validation[$header]['values']),
                            "select" => self::setDropDownValidation($spreadsheet, $activeWorksheet, $header, $validation[$header]['values']),
                            "date" => self::setDateValidation($activeWorksheet, $header, Arr::get($validation[$header], 'before'), Arr::get($validation[$header], 'after')),
                            "time" => self::setTimeValidation($activeWorksheet),
                            "number" => self::setNumberValidation($activeWorksheet, Arr::get($validation[$header], 'min'), Arr::get($validation[$header], 'max')),
                            "text" => Arr::get($validation[$header], 'maxlength') && self::setMaxTextLengthValidation($activeWorksheet, Arr::get($validation[$header], 'maxlength')),
                            "multiple" => self::setMultipleValidation($spreadsheet, Arr::get($validation[$header], 'options')),
                            default => null
                        };

                        $needvalidation = Arr::get($validation[$header], 'required', false);
                        for ($i = 1; $i <= 2000; $i++) {
                            if ($i > 1 && $val != null && is_object($val)) {
                                $activeWorksheet->getCell([$column + 1, $i])->setDataValidation(clone $val);
                            }
                            if($needvalidation) {
                                $activeWorksheet->getStyle([$column + 1, $i])->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('00cc66');
                            }
                            if($borders) {
                                $activeWorksheet->getStyle([$column + 1, $i])->getBorders()->getTop()->setBorderStyle(Border::BORDER_THIN);
                                $activeWorksheet->getStyle([$column + 1, $i])->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THIN);
                                $activeWorksheet->getStyle([$column + 1, $i])->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THIN);
                                $activeWorksheet->getStyle([$column + 1, $i])->getBorders()->getRight()->setBorderStyle(Border::BORDER_THIN);
                            }
                        }
                    }
                }
                if ($filters) {
                    $activeWorksheet->setAutoFilter(
                        $activeWorksheet->calculateWorksheetDimension()
                    );
                }
            } else {
                $index = 0;
                foreach ($headers as $header => $params) {
                    $activeWorksheet->setCellValue([$index + 1, 1], $header);
                    $activeWorksheet->mergeCells([$index + 1, 1, $index + $params[0], $params[1]]);
                    $index += $params[0];
                }
            }
        }
        $row_num = 2;
        foreach ($rows as $row) {
            $column_num = 1;
            $rowparse = $row;
            if (is_object($row)) {
                $rowparse = $row->toArray();
            }

            foreach ($rowparse as $value) {
                if (!is_array($value)) {
                    $activeWorksheet->setCellValue([$column_num, $row_num], $value);

                    $activeWorksheet->getStyle([$column_num, $row_num])->getAlignment()->setWrapText(true);
                } elseif (is_array($value)) {
                    if (array_key_exists('link', $value) && array_key_exists('name', $value)) {
                        if (!is_null($value['link']) && !is_null($value['name'])) {
                            $activeWorksheet->getCell([$column_num, $row_num])->getHyperlink()->setUrl($value['link']);

                            $activeWorksheet->setCellValue([$column_num, $row_num], $value['name']);

                            $activeWorksheet->getStyle([$column_num, $row_num])->getAlignment()->setWrapText(true);
                            $activeWorksheet->getStyle([$column_num, $row_num])->getFont()->setColor(new Color(Color::COLOR_BLUE));
                        }
                    } else
                        if (count($value) == 4) {
                            $cell = $activeWorksheet->setCellValue([$column_num, $row_num], $value[0]);
                            $activeWorksheet->getStyle([$column_num, $row_num])->applyFromArray($value[3]);
                            $activeWorksheet->mergeCells([$column_num, $row_num, $column_num + $value[1] - 1, $row_num + $value[2] - 1]);
                            $column_num += $value[1] - 1;
                        }
                }
                if($borders) {
                    $activeWorksheet->getStyle([$column_num, $row_num])->getBorders()->getTop()->setBorderStyle(Border::BORDER_THIN);
                    $activeWorksheet->getStyle([$column_num, $row_num])->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THIN);
                    $activeWorksheet->getStyle([$column_num, $row_num])->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THIN);
                    $activeWorksheet->getStyle([$column_num, $row_num])->getBorders()->getRight()->setBorderStyle(Border::BORDER_THIN);
                }
                $column_num++;
            }
            $row_num++;
        }

        // self::setDropDownValidation($activeWorksheet, "Компания", [
        //     "Компания1",
        //     "Компания 2"
        // ]);

        // self::setDateValidation($activeWorksheet, "Дата выхода", null, Carbon::now());
        // self::setTimeValidation($activeWorksheet);
        $writer = match ($type) {
            'xlsx' => new Xlsx($spreadsheet),
            'csv' => new Csv($spreadsheet),
            'ods' => new Ods($spreadsheet),
            'html' => new Html($spreadsheet),
        //'pdf' => new Pdf($spreadsheet)
        };

        $temporaryDirectory = (new TemporaryDirectory())->create();
        $path = $temporaryDirectory->path("docCompiled");
        $writer->save("{$path}/table_compiled.{$type}");
        return "{$path}/table_compiled.{$type}";
    }
    /**
     * Summary of wordFile
     * @param string $templateName
     * @param array<mixed> $values
     * @return string
     */
    public static function wordFile(string $templateName, array $values = []): string
    {
        $document = new TemplateProcessor($templateName);
        foreach ($values as $key => $value) {

            if (str_contains($key, 'image_')) {
                //throw new \Exception($key);
                $document->setImageValue($key, $value);
            } elseif (is_array($value)) {
                if (count($value) > 0) {
                    $compiled_values = array_map(fn($value_subojb) => array_combine(
                        array_map(
                            fn($k) => "{$key}_{$k}",
                            array_keys((array) $value_subojb)
                        ),
                        (array) $value_subojb
                    ), $value);
                    if (mb_substr($key, 0, 6) == 'block_') {

                        $document->cloneBlock(
                            $key,
                            count($compiled_values),
                            true,
                            true
                        );
                        for ($i = 1; $i <= count($compiled_values); $i++) {
                            foreach ($compiled_values[$i - 1] as $compile_key => $compile_value) {
                                if (is_object($compile_value)) {
                                    $document->setComplexBlock($compile_key . '#' . $i, $compile_value);
                                } else {
                                    $document->setValue($compile_key . '#' . $i, htmlentities($compile_value));
                                }
                            }
                        }
                    } else {
                        $document->cloneRowAndSetValues(
                            $key . '_' . array_keys((array) $value[0])[0],
                            //$compiled_values
                            array_map(fn($compile_array) => array_map(fn($compile_array_value) => htmlspecialchars($compile_array_value), $compile_array), $compiled_values)
                        );
                    }
                }
            } elseif (is_bool($value)) {
                if ($value == false) {
                    $document->deleteBlock($key);
                }
            } elseif (is_object($value)) {
                $document->setComplexBlock($key, $value);
            } else {
                $document->setValue($key, htmlentities($value));
            }
        }
        $temporaryDirectory = (new TemporaryDirectory())->create();
        $path = $temporaryDirectory->path("docCompiled");
        $document->saveAs("{$path}/compiled.docx");
        return "{$path}/compiled.docx";
    }
    /**
     * Summary of convertDocxToPdf
     * @param string $path
     * @return string
     */
    public static function convertDocxToPdf(string $path): string
    {
        $domPdfPath = base_path('vendor/dompdf/dompdf');
        Settings::setPdfRendererPath($domPdfPath);
        Settings::setPdfRendererName('DomPDF');
        Settings::setPdfRendererOptions([
            'font' => 'DejaVu Sans'
        ]);
        $temporaryDirectory = (new TemporaryDirectory())->create();
        $path_temp = $temporaryDirectory->path("converted");
        $fileFormat = match (pathinfo($path)['extension']) {
            'doc' => 'MsDoc',
            'docx' => 'Word2007',
            'odt' => 'ODText',
            'rtf' => 'RTF',
            'html' => 'HTML',
            default => 'Word2007'
        };

        IOFactory::createWriter(IOFactory::load($path, $fileFormat), 'PDF')->save("{$path_temp}/doc.pdf");

        return "{$path_temp}/doc.pdf";
    }


    public static function archive(array $files, string $zipFileName)
    {
        $zipFileName = "file";
        $zip = new ZipArchive;
        if ($zip->open(public_path("{$zipFileName}.zip"), ZipArchive::CREATE) === TRUE) {
            foreach ($files as $fname => $fpath) {
                $zip->addFile($fpath, $fname);
            }
            if (!$zip->close()) {
                throw new \Exception("Не удалось создать архив");
            }
            return public_path("{$zipFileName}.zip");
        } else {
            throw new \Exception("Не удалось создать архив");
        }

    }


    /**
     * Summary of wordTableCollection
     * @param \Illuminate\Support\Collection $collection
     * @param string $label
     * @return string
     */
    public static function wordTableCollection(Collection $collection, string $label = null): string
    {

        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        if ($label) {
            $text = $section->addText(htmlentities($label));
        }
        $table = $section->addTable('myTable');

        $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
        $temporaryDirectory = (new TemporaryDirectory())->create();
        $path = $temporaryDirectory->path("converted");
        $objWriter->save("{$path}/converted.docx");
        return "{$path}/converted.docx";
    }
}
