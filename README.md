# luler/export-excel

多功能分页将数据导出为excel文件助手

# 助手类列表如下

- MultiPageExcelHelper

# 使用示例

```php

<?php

require 'vendor/autoload.php';
//多页导出时，给每页指定下标即可，如下：
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
//表头带*号会自动标红
\Luler\Excel\MultiPageExcelHelper::exportMultiPageExcelFile('test',
    [
        ['字段1*', '字段2*', '字段3*', '字段4',],
    ]
);

```


