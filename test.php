<?php

require 'vendor/autoload.php';

//\Luler\Excel\MultiPageExcelHelper::exportMultiPageExcelFile('test',
//    [
//        '第一页' => ['字段1*', '字段2*', '字段3*', '字段4',]
//    ],
//    [
//        '第一页' => [
//            [
//                '值1',
//                '值2',
//                '值3',
//                '值4',
//            ]
//        ]
//    ],
//    [
//        '第一页' => ['大标题1'],
//    ]
//);
\Luler\Excel\MultiPageExcelHelper::exportMultiPageExcelFile('test',
    [
        ['字段1*', '字段2*', '字段3*', '字段4',],
    ]
);