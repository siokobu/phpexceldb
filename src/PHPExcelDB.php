<?php

namespace PHPExcelDB;

use Exception;
use PDO;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Monolog\Logger;
use Monolog\Handler\StreamHandler;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use fuga;

/**
 *  Excelを利用したDB操作のユーティリティクラス
 *
 */
class PHPExcelDB {
	
	/** 取り扱うDBの接続 */
	private $pdo;
	
	/** ロガーの生成 */
	private $logger;
	
	/** 定数：DSN用のPostgreSQL */
	public const DSN_PGSQL = "pgsql";
	

	/**
	 * PDO作成用のサポートメソッド．
	 * @param Array $connInfo DB接続情報を指定する．host,port,dbname,username,passwordを設定する．
	 */
	public static function createPDO($connInfo = null)
	{

		$constr = "pgsql:host=".$connInfo["host"].";port=".$connInfo["port"].";dbname=".$connInfo["dbname"].";";
		$username = $connInfo["username"];
		$password = $connInfo["password"];
		
		// DB接続用のPDOオブジェクトを初期化
		$pdo = new PDO($constr, $username, $password);
		$pdo->setAttribute(PDO::ATTR_ERRMODE,PDO::ERRMODE_EXCEPTION);

		return $pdo;
	}
	
	
	/**
	 * コンストラクタ．引数で設定したPDOを設定する．
	 * @param PDO $pdo 設定するPDOオブジェクト
     */
	public function __construct(PDO $pdo) {
		$this->setPdo($pdo);
		$this->logger = new Logger('PHPExcelDB');
		$this->logger->pushHandler(new StreamHandler('php://stdout', Logger::DEBUG));
	}
	
	/**
	 * オブジェクトにPDOを再設定するためのメソッド
	 * @param PDO $pdo
	 */
	public function setPdo(PDO $pdo)
	{
		$this->pdo = $pdo;
	}
	
	/**
	 * Monologによるロガーを設定する．設定しない場合は、標準出力へ出力する標準ロガーを仕様する．
	 * （現段階では、debug出力しかしない）
	 * @param \Monolog\Logger $logger
	 */
	public function setLogger(Logger $logger) {
		$this->logger = $logger;
	}
	
	/**
	 *  Excelファイルの内容をDBへ登録する。
	 *  @param String $inputFile 登録用データファイル
	 *  @param Boolean $isDeleteTable 処理対象のテーブルのデータを削除するかどうか。Trueならば削除する
	 */
	public function importDBFromExcel($inputFile, $isDeleteTable)
	{
		try {
			// 対象のExcelファイルを開く
			$excel = (new \PhpOffice\PhpSpreadsheet\Reader\Xlsx())->load($inputFile);
			
			// 処理対象のテーブル名を全て取得する
			$tables = array();
			for ($i = 0; $i < $excel->getSheetCount(); $i++) {
				array_push($tables, $excel->getSheet($i)->getTitle());
			}
			
			// トランザクションを開始し、処理を開始する
			$this->pdo->beginTransaction();
			
			// 処理対象のテーブルを全て削除する
			if($isDeleteTable) {
				
				// 処理対象のテーブルのデータを削除する
				foreach ($tables as $table) {
					
					// 指定したテーブルのデータをすべて削除する
					$this->deleteTableData($table);
				}
				
			}
	     
			// テーブルの数だけ繰り返し処理を実行する
			foreach ($tables as $tableName) {

				// 処理対象のシートを取得
				$sheet = $excel->getSheetByName($tableName);

				// 最終行を取得する
				$max_row = $sheet->getHighestRow();
				
				// 最終列を取得する
				$max_col = Coordinate::columnIndexFromString($sheet->getHighestColumn())-1;
				
				// メタデータを取得
				$metas = $this->getMetadatas($tableName);
				
				// カラム文字列格納用配列を初期化
				$columns = [];
				
				// 実行するSQL
				$sql = "";
				
				// プリペアドステートメント
				$stmt = null;
				
				// 各行を読み込みながらINSERT
				for ($j = 0; $j < $max_row; $j++) {
					
					// 一行目にはカラム名が設定されている。ここではプリペアドステートメントを生成
					if ($j == 0) {
						$sql = $sql . "INSERT INTO ".$tableName." (";
						
						// INSERT 対象のカラム列を生成
						for ($k = 0; $k < $max_col; $k++) {
							$columns[$k] = $sheet->getCellByColumnAndRow($k+1, $j+1)->getValue();
							$sql = $sql . $columns[$k] . ", ";
						}
						$sql = mb_substr($sql, 0, -2).") VALUES (";
						
						// VALUES句のプレースホルダを作成
						for ($k = 0; $k < $max_col; $k++) {
							$sql = $sql . ":" . $sheet->getCellByColumnAndRow($k+1, $j+1) . ", ";
						}
						$sql = mb_substr($sql, 0, -2).")";
						
						// プリペアドステートメントを生成
						$stmt = $this->pdo->prepare($sql);
						
						// インサート用のプリペアドステートメントを出力
						$this->logger->debug($sql);
						
						continue;
					}
					
					
					// プレースホルダに対して、値を設定する
					for ($k = 0; $k < $max_col; $k++) {
						$value = $sheet->getCellByColumnAndRow($k+1, $j+1)->getValue();
						$stmt->bindValue(":".$columns[$k], $value, $metas[$columns[$k]]["pdo_type"]);
					}
					
					// SQLの実行
					$stmt->execute();
				}
	     	}
	     
	     	// 全部コミット
	     	$this->pdo->commit();
	     } catch (Exception $ex) {
	     	
	     	$this->logger->debug($ex->getMessage());
	     	$this->logger->debug($ex->getTraceAsString());
	     	if($this->pdo->inTransaction()) $this->pdo->rollback();
	     	throw $ex;
	     }
	}
	
	/**
	 * 指定したテーブルの情報をすべて削除する。
	 * @param String $tableName テーブル名
	 */
	private function deleteTableData($tableName) {
		$sql = "DELETE FROM ".$tableName.";";
		$this->logger->debug($tableName."テーブルを削除します：".$sql);
		$this->pdo->exec($sql);
	}
	

	/**
	 * DBの内容をExcelへExportする
	 * @param String $outputFile 出力用データファイル
	 * @param Array $targetTables 対象テーブル
	 */
	public function exportDBtoExcel($outputFile, $targetTables)
	{
		try{
			// SpreadSheetオブジェクトを初期化し、デフォルトのシートを削除
			$excel = new Spreadsheet();
			$excel->removeSheetByIndex(0);
			
			// 設定されたテーブルごとに出力処理を行う
			for ($i = 0; $i < count($targetTables); $i++ ) {
				
				// 処理対象のシートを作成
				$sheet = $excel->createSheet($i);
				
				// 作成したシートにタイトルを設定
				$sheet->setTitle($targetTables[$i]);
				
				// カラム一覧を取得するために、SELECT文を投入する
				$stmt = $this->pdo->query("SELECT * FROM $targetTables[$i]");
				
				// メタデータを取得
				$meta = array();
				for($j = 0; $j < $stmt->columnCount(); $j++) {
					$meta[$j] = $stmt->getColumnMeta($j);
				}
				
				// 一行目を出力
				for ($j = 0; $j < count($meta); $j++) {
					$sheet->setCellValueByColumnAndRow($j+1, 1, $meta[$j]['name'], );
				}
				
				$j = 2;
				while($result = $stmt->fetch(PDO::FETCH_NUM)) {
					for($k = 0; $k < count($result); $k++) {
						$sheet->getCellByColumnAndRow($k+1, $j)->getStyle()->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_TEXT);
						$sheet->setCellValueByColumnAndRow($k+1, $j, $result[$k]);
					}
					$j++;
				}
				
			}
			$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($excel);
			$writer->save($outputFile);
			
			
		} catch (Exception $ex) {
			$this->logger->error($ex->getMessage());
			$this->logger->error($ex->getTraceAsString());
			if($this->pdo->inTransaction()) $this->pdo->rollback();
			throw $ex;
		}
	}
	
	/**
	 * 二つのExcelファイルの差分を取得する．このメソッドに限り、PDOを利用しない．
	 * @param String $file1 差分取得元のパス
	 * @param String $file2 差分取得元のパス
	 * @param String $diff 差分出力ファイル
	 * @return boolean 差分があったかどうか．差分がない場合、trueを返す．
	 */
	public function getExcelDiff($file1, $file2, $diff) {
		// 戻り値．差分がある場合、falseとする．
		$returnValue = false;
		
		$diffFile = new Spreadsheet();
		$diffFile->removeSheetByIndex(0);
		$bookFile1 = (new \PhpOffice\PhpSpreadsheet\Reader\Xlsx())->load($file1);
		$bookFile2 = (new \PhpOffice\PhpSpreadsheet\Reader\Xlsx())->load($file2);
		
		// file1のシート名をすべて取得
		$file1SheetNames = $bookFile1->getSheetNames();
		for($i = 0; $i < count($file1SheetNames); $i++) $this->logger->debug("file1:".$file1SheetNames[$i]);

		// file2のシート名をすべて取得
		$file2SheetNames = $bookFile2->getSheetNames();
		for($i = 0; $i < count($file2SheetNames); $i++) $this->logger->debug("file2:".$file2SheetNames[$i]);
		
		// 調査対象のシート名を配列で取得
		$targetSheets = [];
		for ($i = 0; $i < count($file1SheetNames); $i++) {
			for ($j = 0; $j < count($file2SheetNames); $j++) {
				if ($file1SheetNames[$i] == $file2SheetNames[$j]) {
					$targetSheets[count($targetSheets)] = $file1SheetNames[$i];
				}
			}
		}
		for($i = 0; $i < count($targetSheets); $i++) $this->logger->debug("targetSheets:".$targetSheets[$i]);

		// file1のみのシートを取得
		$file1Only = [];
		for ($i = 0; $i < count($file1SheetNames); $i++) {
			if(!in_array($file1SheetNames[$i], $targetSheets)) $file1Only[count($file1Only)] = $file1SheetNames[$i];
		}
		for($i = 0; $i < count($file1Only); $i++) $this->logger->debug("file1Only:".$file1Only[$i]);
		
		// file2のみのシートを取得
		$file2Only = [];
		for ($i = 0; $i < count($file2SheetNames); $i++) {
			if(!in_array($file2SheetNames[$i], $targetSheets)) $file2Only[count($file2Only)] = $file2SheetNames[$i];
		}
		for($i = 0; $i < count($file2Only); $i++) $this->logger->debug("file2Only:".$file2Only[$i]);
		
		// 対象のシートごとに差分を調査
		for ($i = 0; $i < count($targetSheets); $i++) {
			// シート行列の最大値を取得、sheet1,sheet2の大きいほうを取得
			$sheet1 = $bookFile1->getSheetByName($targetSheets[$i]);
			$sheet2 = $bookFile2->getSheetByName($targetSheets[$i]);
			$maxCell1 = $sheet1->getHighestRowAndColumn();
			$maxCell2 = $sheet2->getHighestRowAndColumn();
			$maxCell = [];
			$maxCell1["row"] > $maxCell2["row"] ? $maxCell["row"] = $maxCell1["row"] : $maxCell["row"] = $maxCell2["row"];
			$maxCell1["column"] > $maxCell2["column"] ? $tmpMaxColumn = $maxCell1["column"] : $tmpMaxColumn = $maxCell2["column"];
			$maxCell["column"] = Coordinate::columnIndexFromString($tmpMaxColumn);
			$this->logger->debug($targetSheets[$i].":maxrow:".$maxCell["row"]);
			$this->logger->debug($targetSheets[$i].":maxcolumn:".$maxCell["column"]);
			
			// 差分登録用のシートを作成
			$sheet = $diffFile->createSheet($i);
			$sheet->setTitle($sheet1->getTitle());
			
			// 各セルを比較し、Diffファイルを作成
			for($j = 0; $j < $maxCell["column"]; $j++) {
				for ($k = 0; $k < $maxCell["row"]; $k++) {
					$cell = Coordinate::stringFromColumnIndex($j+1).($k+1);
					$value1 = $sheet1->getCell($cell)->getValue();
					$value2 = $sheet2->getCell($cell)->getValue();
					// 値が同じ場合は値をそのまま書き込む
					if($value1 == $value2) {
						$sheet->setCellValue($cell, $value1);
					// 値が異なる場合は、"::"区切りで両方を書き込む
					} else {
						$sheet->setCellValue($cell, $value1."::".$value2);
						$sheet->getStyle($cell)->applyFromArray(['fill' => ['fillType' => Fill::FILL_SOLID, 'color' => ['argb' => COLOR::COLOR_RED]]]);
					    $returnValue = false;
					}
				}
			}
		}
		
		// 差分ファイルの書き出し
		$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($diffFile);
		$writer->save($diff);
		
		return $returnValue;
		
	}
	
	/**
	 * 
	 * @param String $tableName
	 * @return array[] カラム名 => メタデータオブジェクト という形式の連想配列
	 */
	private function getMetadatas($tableName) {
		// メタデータ格納用の配列
		$metas = array();
		
		// メタデータ取得用のSELECTを実行する
		$stmt = $this->pdo->query("SELECT * FROM ".$tableName);
		$count = $stmt->columnCount();
		if(!$count) print_r($count);
		// 
		for($i = 0; $i < $count; $i++) {
			
			// 結果セットからメタデータを取得
			$meta = $stmt->getColumnMeta($i);
			
			// メタデータを配列に格納
			$metas[$meta["name"]] = $meta;
//			array_merge($metas, [$meta["name"] => $meta]);
		}
		return $metas;
	}

	
}