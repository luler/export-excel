<?php
/**
 * Created by PhpStorm.
 * @version : 1.0
 * User: Alan_
 * Date: 2017/8/8
 * Time: 10:49
 */

namespace Luler\Excel;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;

class MultiPageExcelHelper
{
    /**
     * 导出excel（多页）
     * @param $file_name //文件名
     * @param $header_titles //列头标题 （二维数组）['第一页'=>['列头一']]
     * @param $datas //数据  （三维数组）['第一页'=>[['第一行数据']]]
     * @param $banners //大标题  （空数组或二维数组）['第一页'=>['大标题']]
     * @param array $widths //宽度设置
     * @param int $height //行高度设置
     * @param bool $is_auto_wrap //是否自动分行
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @author 我只想看看蓝天 <1207032539@qq.com>
     */
    public static function exportMultiPageExcel(string $file_name, array $header_titles, array $datas, array $banners = [], array $widths = [], int $height = null, $is_auto_wrap = false)
    {
        $php_excel = self::buildSheet($banners, $header_titles, $datas, $widths, $height);
        //设置excel导出
        $writer = IOFactory::createWriter($php_excel, 'Xlsx');

        //中文名兼容各种浏览器
        $ua = $_SERVER["HTTP_USER_AGENT"];
        if (preg_match("/MSIE/", $ua)) {
            header('Content-Disposition: attachment; filename="' . $file_name . '.xlsx"');
        } else if (preg_match("/Firefox/", $ua)) {
            header('Content-Disposition: attachment; filename*="utf8\'\'' . $file_name . '.xlsx"');
        } else {
            header('Content-Disposition: attachment; filename=' . urlencode($file_name . '.xlsx'));
        }

        header('Cache-Control: max-age=0');
        header('Content-Type:application/vnd.ms-excel');

        $writer->save("php://output");
    }

    /**
     * 导出excel文件(多页)
     * @param $file_path //文件路径
     * @param $header_titles //列头标题 （二维数组）['第一页'=>['列头一']]
     * @param $datas //数据  （三维数组）['第一页'=>[['第一行数据']]]
     * @param $banners //大标题  （空数组或二维数组）['第一页'=>['大标题']]
     * @param array $widths //宽度设置
     * @param int $height //行高度设置
     * @param bool $is_auto_wrap //是否自动分行
     * @return bool
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @author 我只想看看蓝天 <1207032539@qq.com>
     */
    public static function exportMultiPageExcelFile($file_path, array $header_titles, array $datas, array $banners = [], array $widths = [], int $height = null, $is_auto_wrap = false)
    {
        $php_excel = self::buildSheet($header_titles, $datas, $banners, $widths, $height);
        //设置excel导出
        $writer = IOFactory::createWriter($php_excel, 'Xlsx');

        $writer->save($file_path . '.xlsx');

        return true;
    }

    /**
     * 构建工作表对象
     * @param $header_titles //列头标题 （二维数组）['第一页'=>['列头一']]
     * @param $datas //数据  （三维数组）['第一页'=>[['第一行数据']]]
     * @param $banners //大标题  （空数组或二维数组）['第一页'=>['大标题']]
     * @param array $widths //宽度设置
     * @param int $height //行高度设置
     * @param bool $is_auto_wrap //是否自动分行
     * @return Spreadsheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @author 我只想看看蓝天 <1207032539@qq.com>
     */
    private static function buildSheet(array $header_titles, array $datas, array $banners = [], array $widths = [], int $height = null, $is_auto_wrap = false)
    {
        $php_excel = new Spreadsheet();
        $index = 0;
        foreach ($header_titles as $key => $titles) {
            if (empty($titles)) {
                throw new \Exception('请设置excel的标题头');
            }
            if ($index == 0) {
                $sheet = $php_excel->getActiveSheet()->setTitle($key);
            } else {
                $sheet = $php_excel->createSheet($index)->setTitle($key);
            }
            $index++; //页数添加
            $data_rows = count($datas[$key]);
            $banner_rows = count($banners[$key] ?? []);
            $data_row_start = $banner_rows + 1;//第几行开始写数据
            $header_titles_columns = count($titles);
            $header_arr = self::getColumnAlphabetRange($header_titles_columns + 5); //增加5个列范围，防止溢出
            //插大标题
            for ($i = 0; $i < $banner_rows; $i++) {
                $sheet->mergeCells('A' . ($i + 1) . ':' . $header_arr[$header_titles_columns - 1] . ($i + 1));
                //设置字体大小
                $sheet->getStyle('A' . ($i + 1))->getFont()->setSize(16);
                //设置粗体
                $sheet->getStyle('A' . ($i + 1))->getFont()->setBold(true);
                $sheet->setCellValue('A' . ($i + 1), $banners[$key][$i]);
            }
            //插表头标题
            for ($i = 0; $i < $header_titles_columns; $i++) {
                //设置宽度
                if (isset($widths[$i]) && is_numeric($widths[$i])) {
                    $sheet->getColumnDimension($header_arr[$i])->setWidth($widths[$i]);
                }
                //设置颜色
                if (strpos($titles[$i], '*') !== false) {
                    $sheet->getStyle($header_arr[$i] . ($banner_rows + 1))->getFont()->getColor()->setARGB(Color::COLOR_RED);
                }
                $sheet->setCellValue($header_arr[$i] . ($banner_rows + 1), $titles[$i]);
            }
            //插数据项
            for ($i = 0; $i < $data_rows; $i++) {
                for ($j = 0; $j < count($datas[$key][$i]); $j++) {
                    $sheet->setCellValueExplicit($header_arr[$j] . ($i + $data_row_start + 1), $datas[$key][$i][$j], DataType::TYPE_STRING);
                }
            }
            //设置行高
            if (is_numeric($height)) {
                for ($i = 0; $i < ($banner_rows + $data_rows + 1); $i++) {
                    $sheet->getRowDimension(($i + 1))->setRowHeight($height);
                }
            }
            //设置居中
            $sheet->getStyle('A1:' . $header_arr[$header_titles_columns - 1] . ($banner_rows + $data_rows + 1))->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            //所有垂直居中
            $sheet->getStyle('A1:' . $header_arr[$header_titles_columns - 1] . ($banner_rows + $data_rows + 1))->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            //所有自动换行
            $is_auto_wrap && $sheet->getStyle('A1:' . $header_arr[$header_titles_columns - 1] . ($banner_rows + $data_rows + 1))->getAlignment()->setWrapText(true);
        }

        return $php_excel;
    }

    /**
     * 获取列字母标志范围
     * @param $count
     * @return array
     * @author 我只想看看蓝天 <1207032539@qq.com>
     */
    private static function getColumnAlphabetRange($count)
    {
        $getColumnAlphabet = function ($count) {
            $res = [];
            $from = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
            $mark = $count >= 1;
            while ($mark) {
                $index = $count % 26;
                $count = $count / 26;
                if ($index == 0 && is_int($count)) {
                    $index = 26;
                }
                if (!empty($res) && $res[0] == 'Z') {
                    $index--;
                }
                array_unshift($res, $from[$index - 1]);
                if ($count <= 1) {
                    $mark = false;
                }
                $count = floor($count);
            }

            return join($res);
        };

        $range = [];
        for ($i = 1; $i <= $count; $i++) {
            $range[] = $getColumnAlphabet($i);
        }

        return $range;
    }
}
