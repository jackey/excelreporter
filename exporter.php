<?php

require 'libraries/phpexcel/Classes/PHPExcel.php';

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

//初始化一个Excel对象
function init_excel() {
	$objPHPExcel = new PHPExcel();

	$objPHPExcel->getProperties()->setCreator("Masterprint")
								 ->setLastModifiedBy("Masterprint")
								 ->setTitle("Product Apply Document")
								 ->setSubject("Product Apply Document")
								 ->setDescription("")
								 ->setKeywords("")
								 ->setCategory("");
	return $objPHPExcel;
}

function init_excel_header($excel, $headers) {
	$sheet = $excel->setActiveSheetIndex(0);
	foreach ($headers as $index => $header) {
		$char_value = ord('A');
		$char_char = chr($char_value + $index);
		$sheet->setCellValue($char_char.'1', $header);
	}
}

function excel_write_one_row($excel, $data, $row_index) {
	$sheet = $excel->setActiveSheetIndex(0);

	foreach ($data as $index => $column_value) {
		$char_value = ord('A');
		$char_char = chr($char_value + $index);
		$sheet->setCellValue($char_char.($row_index + 1), $column_value);
	}
}

function excel_rename_sheet($excel, $sheet_index, $name) {
	$excel->setActiveSheetIndex($sheet_index);
	$excel->getActiveSheet()->setTitle($name);
}

function save_excel($excel, $name) {
	// Save Excel 2007 file
	$filename = $name. '.xlsx';
	$objWriter = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
	$objWriter->save($filename);
	return getcwd().'/'.$filename;
}

function process_apply_time($time) {
	return date('Y-d-m', floatval($time));
}

function process_approve_time($time) {
	if ($time == 0) {
		return '在审核中';
	}
	return date('Y-d-m', floatval($time));
}

//$file should be absolute path.
function mail_attachment($to, $subject, $message, $from, $file) {
  	//$file should include path and filename
	$filename = basename($file);
	$file_size = filesize($file);
	$content = chunk_split(base64_encode(file_get_contents($file)));
	$uid = md5(uniqid(time()));
	$from = str_replace(array("\r", "\n"), '', $from); // to prevent email injection
	$header = "From: Masterprint<".$from.">\r\n"
		."MIME-Version: 1.0\r\n"
		."Content-Type: multipart/mixed; boundary=\"".$uid."\"\r\n\r\n"
		."This is a multi-part message in MIME format.\r\n" 
		."--".$uid."\r\n"
		."Content-type:text/plain; charset=iso-8859-1\r\n"
		."Content-Transfer-Encoding: 7bit\r\n\r\n"
		.$message."\r\n\r\n"
		."--".$uid."\r\n"
		."Content-Type: application/octet-stream; name=\"".$filename."\"\r\n"
		."Content-Transfer-Encoding: base64\r\n"
		."Content-Disposition: attachment; filename=\"".$filename."\"\r\n\r\n"
		.$content."\r\n\r\n"
		."--".$uid."--";
	return mail($to, $subject, "", $header);
}

function load_user_email($uid) {
	global $mysql;

	$query = mysql_query("SELECT * FROM users where uid = ". $uid);
	$row = mysql_fetch_assoc($query);
	if ($row)
		return $row['mail'];
	else 
		return FALSE;
}

function process_user_report($uid) {
	global $config;
	$mail = load_user_email($uid);
	if (!$mail) {
		return FALSE;
	}
	$sheets = $config['sheets'];

	foreach ($sheets as $index => $sheet) {
		if (empty($sheet)) {
			continue;
		}
		$sql = $sheet['sql'];
		$sql .= " WHERE product_apply.uid = ". $uid;
		$columns = $sheet['columns'];
		$name = $sheet['name'];
		$headers = array_values($columns);
		//初始化excel对象
		$excel = init_excel();
		//设置Excel第一行
		init_excel_header($excel, $headers);
		$result = mysql_query($sql);

		if (!$result) {
			die("MYSQL error". mysql_error());
		}
		$crt_index = 1;
		while ($row = mysql_fetch_assoc($result)) {
			$data = array();
			foreach ($columns as $column_name => $column) {
				$column_value = $row[$column_name];
				if (!$column_value) {
					$column_value = "无信息记录";
				}
				if (function_exists('process_'.$column_name)) {
					$column_value = call_user_func('process_'.$column_name, $column_value);
				}
				$data[] = $column_value;
			}
			excel_write_one_row($excel, $data, $crt_index);
			$crt_index += 1;
		}
		excel_rename_sheet($excel, $index, $name);
		$path = save_excel($excel, 'product_apply'. $uid);
		//$to = 'v-beche@microsoft.com';
		$to = 'jziwenchen@gmail.com';
		$from = '397420507@qq.com';
		$message = "您好,\n附件是您在Masterprint 系统下印刷产品的使用统计数据.\n谢谢\n";
		mail_attachment($to, '产品的印刷统计数据', wordwrap($message), $from, $path);
	}
}

$config = require('config.php');

$db = $config['db'];

$mysql = mysql_connect($db['host'], $db['user'], $db['password']);

mysql_select_db($db['database'], $mysql);

$users = array();
$query_users = mysql_query("SELECT distinct uid FROM product_apply");

while ($row = mysql_fetch_assoc($query_users)) {
	$users[] = $row['uid'];
}

foreach ($users as $uid) {
	process_user_report($uid);
}

mysql_close($mysql);
