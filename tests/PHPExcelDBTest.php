<?php

namespace PHPExcelDBTest;

use PHPUnit\Framework\TestCase;
use PHPExcelDB\PHPExcelDB;
use PDO;

class PHPExcelDBTest extends TestCase {
	
	private const DSN = "pgsql:host=127.0.0.1;port=5432;dbname=pgsqltest";
	private const USERNAME = "tomocky1";
	private const PASSWORD = "siokobu8400";
	private const INPUTDIR = __DIR__."/input/";
	private const OUTPUTDIR = __DIR__."/output/";
	private const DIFFDIR = __DIR__."/diff/";
	
	private $pdo;
	
	protected function setUp() {
		$this->pdo = new PDO(self::DSN, self::USERNAME, self::PASSWORD);
		$this->pdo->exec(file_get_contents(__DIR__."/sql/PgSQLTest.sql"));
	}
	
	/**
	 * @test
	 * @requires PHP 7.4
	 */
	public function test01() {
		$phpExcelDB = new PHPExcelDB($this->pdo);
		$phpExcelDB->importDBFromExcel(self::INPUTDIR."test01_01.xlsx");
		
		$phpExcelDB->exportDBtoExcel(self::OUTPUTDIR."test01_01.xlsx", ["parent_table", "child_table"]);
		
		$this->assertTrue(true);
		
	}
	
	/**
	 * @test
	 */
	public function testGetExcelDiff_01() {
		$phpExcelDB = new PHPExcelDB($this->pdo);
		$phpExcelDB->getExcelDiff(self::INPUTDIR."testGetExcelDiff_01_01.xlsx", self::INPUTDIR."testGetExcelDiff_01_02.xlsx", self::DIFFDIR."testGetExcelDiff_01.xlsx");
		
	}
}