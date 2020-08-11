<?php
namespace app\commands;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\SheetView;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use yii\console\Controller;
use yii\console\ExitCode;
use yii\db\Exception;
use yii\helpers\ArrayHelper;
use yii\helpers\Console;

/**
 * Excel Tools and Utilities
 * @package app\commands
 */
class ExcelController extends Controller
{
    protected $supportedExtensions = ['xlsx', 'xls'];
    public $inputDir = '';
    public $outputDir = '';
    public $outputFilename = '';
    public $excelFiles = [];
    public $keyColumn = '';
    public $data;

    /**
     * Merge excel files into one
     * @param string $inputDir
     * @param string $outputDir
     * @param string $keyColumn
     * @param string $outputFilename
     * @return int
     */
    public function actionMerge($inputDir = '', $outputDir = '', $keyColumn = '', $outputFilename = '')
    {
        try{
            $this->inputDir = $inputDir;
            while(!is_dir($this->inputDir) || !is_readable($this->inputDir)){
                $this->inputDir = Console::input('Input directory: ', Console::FG_YELLOW);
                if(!is_readable($this->inputDir)){
                    Console::stdout("{$this->inputDir} is not readable\n", Console::FG_RED);
                    $this->inputDir = '';
                }
            }

            $this->outputDir = $outputDir;
            while(!is_dir($this->outputDir) || !is_writable($this->outputDir)){
                $this->outputDir = Console::input('Output file: ', Console::FG_YELLOW);
                if(!is_writable($this->outputDir)){
                    Console::stdout("{$this->outputDir} is not writable\n", Console::FG_RED);
                    $this->outputDir = '';
                }
            }

            $this->keyColumn = $keyColumn;
            while($this->keyColumn == ''){
                $this->keyColumn = Console::input('Key column: ', Console::FG_YELLOW);
            }

            $this->outputFilename = $outputFilename;
            $isValidOutputFileName = false;
            while(!$isValidOutputFileName){
                if($this->outputFilename == ''){
                    $this->outputFilename = Console::input('Output filename: ', Console::FG_YELLOW);
                }
                if(!in_array(pathinfo($this->outputDir.'/'.$this->outputFilename, PATHINFO_EXTENSION), $this->supportedExtensions)){
                    Console::stdout("File extension is not supported!\n", Console::FG_RED);
                    $this->outputFilename = '';
                }
                else if(is_file($this->outputDir.'/'.$this->outputFilename)){
                    if(Console::input("File already exists! Overwrite? (Y/n): ", Console::FG_YELLOW) == 'Y'){
                        unlink($this->outputDir.'/'.$this->outputFilename);
                    } else {
                        $this->outputFilename = '';
                    }
                }
                else {
                    $isValidOutputFileName = true;
                }
            }

            foreach($this->supportedExtensions as $extension){
                foreach (glob($this->inputDir."/*.{$extension}") as $filename) {
                    $this->excelFiles[] = $filename;
                }
            }

            if(empty($this->excelFiles)){
                Console::stdout("No file found!\n", Console::FG_RED);
                return ExitCode::OK;
            } else {
                //sort by name
                sort($this->excelFiles);
                Console::output('------------------------------');
                foreach($this->excelFiles as $filename){
                    Console::stdout($filename ."\n", Console::FG_GREEN);
                }
                Console::output('------------------------------');
            }

            if(Console::input("Are you sure you want to continue (Y/n): ", Console::FG_YELLOW) == 'Y'){
                $this->readInputFiles();
                $this->writeOutputFile();
            }

        } catch (\Exception $ex){
            Console::output('Error: '.$ex->getMessage());
        }

        return ExitCode::OK;
    }

    private function readInputFiles()
    {
        $totalColumns = 0;
        foreach($this->excelFiles as $filename) {
            try {
                $reader = IOFactory::createReaderForFile($filename);
                $reader->setReadDataOnly(true);
                $reader->setLoadAllSheets();

                $spreadsheet = $reader->load($filename);


                $worksheetNames = $reader->listWorksheetInfo($filename);
                foreach ($worksheetNames as $worksheet) {
                    //$this->data[$filename][$worksheet['worksheetName']] = $worksheet;
                }

                $currentWorksheet = $spreadsheet->getActiveSheet();
                $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);
                $keyColumnIndex = '';

                foreach ($sheetData[1] as $columnIndex => $cellData) {
                    if ($cellData == $this->keyColumn) {
                        $keyColumnIndex = $columnIndex;
                    }
                }
                if ($keyColumnIndex == '') {
                    Console::stdout("Error reading file {$filename}: Key column not found!\n", Console::FG_RED);
                    return;
                } else {
                    $sheetData = ArrayHelper::index($sheetData, $keyColumnIndex);

                    $colIndexMapping = [];
                    foreach(array_keys($sheetData[$this->keyColumn]) as $colIndex){
                        if($colIndex != $keyColumnIndex){
                            $totalColumns++;
                            $colIndexMapping[$colIndex] = $totalColumns;
                        } else {
                            $colIndexMapping[$colIndex] = $colIndex;
                        }
                    }

                    //change column index
                    foreach($sheetData as $rowIndex => $row){
                        foreach($row as $colIndex => $cellData){
                            $this->data[$filename][$rowIndex][$colIndexMapping[$colIndex]] = $cellData;
                        }
                    }
                }
            } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
                Console::stdout('Error loading file: ' . $e->getMessage(), Console::FG_RED);
            }
        }
    }

    private function writeOutputFile()
    {
        if(empty($this->data)){
            throw new Exception('Empty output data!');
        }

        $outputData = [];
        foreach($this->data as $file => $rows){
            foreach($rows as $rowIndex => $row){
                if(isset($outputData[$rowIndex])){
                    $outputData[$rowIndex] = $outputData[$rowIndex] + $row;
                } else {
                    $outputData[$rowIndex] = $row;
                }
            }
        }

        $outputFile = $this->outputDir.'/'.$this->outputFilename;

        try {
            $spreadsheet = new Spreadsheet();
            $worksheet = $spreadsheet->getActiveSheet();
            $worksheet->fromArray($outputData);

            $writer = IOFactory::createWriter($spreadsheet, ucfirst(pathinfo($this->outputDir.'/'.$this->outputFilename, PATHINFO_EXTENSION)));
            //$writer->setPreCalculateFormulas(false);
            $writer->save($outputFile);


        } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
            Console::stdout("Error writing file {$outputFile}: " . $e->getMessage(), Console::FG_RED);
        }
    }
}