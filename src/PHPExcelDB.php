<?php

namespace PHPExcelDB;

use Exception;
use PDO;
use PhpOffice\PhpSpreadsheet\Spreadsheet;


/**
 *  
 *
 */
class PHPExcelDB {
	
	/** 取り扱うDBの接続 */
	private $pdo;
	
	/**
	 * 
     */
	public function __construct($pdo) {
		$this->pdo = $pdo;
	}
	
	/**
	 *  
	 */
	public function importDBFromExcel($inputFile)
	{
		try {
			$excel = (new \PhpOffice\PhpSpreadsheet\Reader\Xlsx())->load($inputFile);
			
			
			$this->pdo->beginTransaction();
	     
			
			for ($i = 0; $i < $excel->getSheetCount(); $i++)
			{
				$sheet = $excel->getSheet($i);
				$tableName = $sheet->getTitle();

				$sql = "DELETE FROM ".$tableName.";";
				print $sql."\n";
				$this->pdo->exec($sql);
						
				$data = $sheet->toArray();
				for ($j = 0; $j < count($data); $j++) {
					// 一行目はコメントとなる
					if ($j == 0) {
						continue;
					}
	         		$sql = "";
	         		$sql .= "INSERT INTO ".$tableName." VALUES (";
	         		for ($k = 0; $k < count($data[$j]); $k++) {
	         			$sql .= "'".$data[$j][$k]."', ";
	         		}
	         		$sql = mb_substr($sql, 0, mb_strlen($sql)-6);
	         		$sql .= ");";
	         		print $sql."\n";
					$this->pdo->exec($sql);

	         	}
	     	}
	     
	     	$this->pdo->commit();
	     } catch (Exception $ex) {
	     	print $ex->getMessage();
	     	print $ex->getTraceAsString();
	     	$this->pdo->rollback();
	     }
	}

	public function exportDBtoExcel($outputFile, $targetTables)
	{
		try{
			$excel = new Spreadsheet();
			
			for ($i = 0; $i < count($targetTables); $i++ ) {
				$sheet = $excel->createSheet($i);
				$sheet->setTitle($targetTables[$i]);
				
				$stmt = $this->pdo->query("SELECT * FROM $targetTables[$i]");
				for($j = 0; $j < $stmt->columnCount(); $j++) {
					$meta = $stmt->getColumnMeta($j);
					print $meta['name']."\n";
					$sheet->setCellValueByColumnAndRow($j+1, 1, $meta['name']);
				}
				
				$j = 2;
				while($result = $stmt->fetch(PDO::FETCH_NUM)) {
					for($k = 0; $k < count($result); $k++) {
						$sheet->setCellValueByColumnAndRow($k+1, $j, $result[$k]);
					}
					$j++;
				}
				
			}
			$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($excel);
			$writer->save($outputFile);
			
			
			
		} catch (Exception $ex) {
			print $ex->getMessage();
			print $ex->getTraceAsString();
			$this->pdo->rollback();
		}
	}

	
}