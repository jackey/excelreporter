<?php

$config = array(
	'sheets' => array(
		array(
			'sql' => 'SELECT * FROM `product_apply` LEFT JOIN users ON users.uid = product_apply.uid LEFT JOIN node on node.nid = product_apply.product_id',
			'columns' => array(
				'name' => '用户名',
				'title' => '产品名称', 
				'apply_time' => '申请时间', 
				'approve_time' => '审核时间', 
				'count' => '申请个数',
				'address' => '发货地址',
				'delivery_company' => '物流公司',
				'delivery_serial_no' => '运单号',
			),
			'name' => '产品印刷申请统计',
		),
		array(
			// 'sql' => "SELECT * FROM node LEFT JOIN users on users.uid = node.uid",
			// 'columns' => array(
			// 	'name' => '用户名',
			// 	'title' => '产品名称',
			// 	''
			// ),
		),
	),
	'db' => array(
		'host' => 'localhost',
		'port' => '3306',
		'user' => 'root',
		'password' => 'admin',
		'database' => 'master_print',
	),
);

return $config;
