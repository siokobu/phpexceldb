<?php

namespace PHPExcelDB;

use Exception;
use RuntimeException;

class PHPExcelDBException extends RuntimeException {

	public function __construct($message, $code = 0, Exception $previous = null) {
		parent::__construct($message, $code, $previous);
	}
	
}