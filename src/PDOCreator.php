<?php

namespace PHPExcelDB;

use Dotenv\Dotenv;
use PDO;

class PDOCreator {
	
	public static function getPDO($constr = null, $username = null, $password = null)
	{
		
		// $constr が設定されていない場合は、.envから接続文字列を取得
		if ($constr == null) {
			$dotenv = Dotenv::create(__DIR__.'/..');
			$dotenv->load();
			
			$host = getenv("DB_HOST");
			$port = getenv("DB_PORT");
			$database = getenv("DB_DATABASE");
			
			$constr = 'pgsql:host='.$host.':'.$port.';dbname='.$database;
		}
		
		// $username が設定されていない場合は、.envから接続文字列を取得
		$username == null ? $username = getenv("DB_USERNAME") : null;
		
		// $password が設定されていない場合は、.envからパスワードを取得
		$password == null ? $password = getenv("DB_PASSWORD") : null;
		
		// DB接続用のPDOオブジェクトを初期化
		print "CONSTR = ".$constr."\n";
		$pdo = new PDO($constr, $username, $password);
		$pdo->setAttribute(PDO::ATTR_ERRMODE,PDO::ERRMODE_EXCEPTION);
		
		return $pdo;
	}
}