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

class MultiPageExcelHelper
{
    /**
     * 导出excel（多页）
     * @param $file_name //文件名
     * @param $banners //大标题  （空数组或二维数组）['第一页'=>['大标题']]
     * @param $header_titles //列头标题 （二维数组）['第一页'=>['列头一']]
     * @param $datas //数据  （三维数组）['第一页'=>[['第一行数据']]]
     * @param array $widths //宽度设置
     * @param int $height //行高度设置
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @author LinZhou <1207032539@qq.com>
     */
    public static function exportMultiPageExcel(string $file_name, array $banners, array $header_titles, array $datas, array $widths = [], int $height = 30)
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
            $header_arr = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
            $header_arr = array_merge($header_arr, ['AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ']);
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
                $sheet->getColumnDimension($header_arr[$i])->setWidth($widths[$i] ?? 24);
                //设置粗体
//            $sheet->getStyle(chr(ord('A') + $i) . ($banner_rows + 1))->getFont()->setBold(true);
                $sheet->setCellValue($header_arr[$i] . ($banner_rows + 1), $titles[$i]);
            }
            //插数据项
            for ($i = 0; $i < $data_rows; $i++) {
                for ($j = 0; $j < count($datas[$key][$i]); $j++) {
                    $sheet->setCellValueExplicit($header_arr[$j] . ($i + $data_row_start + 1), $datas[$key][$i][$j], DataType::TYPE_STRING);
                }
            }
            //设置行高
            for ($i = 0; $i < ($banner_rows + $data_rows + 1); $i++) {
                $sheet->getRowDimension(($i + 1))->setRowHeight($height);
            }
            //设置居中
            $sheet->getStyle('A1:' . $header_arr[$header_titles_columns - 1] . ($banner_rows + $data_rows + 1))->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            //所有垂直居中
            $sheet->getStyle('A1:' . $header_arr[$header_titles_columns - 1] . ($banner_rows + $data_rows + 1))->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            //所有自动换行
            $sheet->getStyle('A1:' . $header_arr[$header_titles_columns - 1] . ($banner_rows + $data_rows + 1))->getAlignment()->setWrapText(true);
        }

        //设置excel导出
        $writer = IOFactory::createWriter($php_excel, 'Xls');

        //中文名兼容各种浏览器
        $ua = $_SERVER["HTTP_USER_AGENT"];
        if (preg_match("/MSIE/", $ua)) {
            header('Content-Disposition: attachment; filename="' . $file_name . '.xls"');
        } else if (preg_match("/Firefox/", $ua)) {
            header('Content-Disposition: attachment; filename*="utf8\'\'' . $file_name . '.xls"');
        } else {
            header('Content-Disposition: attachment; filename=' . urlencode($file_name . '.xls'));
        }

        header('Cache-Control: max-age=0');
        header('Content-Type:application/vnd.ms-excel');

        $writer->save("php://output");
    }

    /**
     * 导出excel文件(多页)
     * @param $file_name //文件名
     * @param $banners //大标题  （空数组或二维数组）['第一页'=>['大标题']]
     * @param $header_titles //列头标题 （二维数组）['第一页'=>['列头一']]
     * @param $datas //数据  （三维数组）['第一页'=>[['第一行数据']]]
     * @param array $widths //宽度设置
     * @param int $height //行高度设置
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @author LinZhou <1207032539@qq.com>
     */
    public static function exportMultiPageExcelFile($file_path, array $banners, array $header_titles, array $datas, array $widths = [], int $height = 30)
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
            $header_arr = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
            $header_arr = array_merge($header_arr, ['AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ']);
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
                $sheet->getColumnDimension($header_arr[$i])->setWidth($widths[$i] ?? 24);
                //设置粗体
//            $sheet->getStyle(chr(ord('A') + $i) . ($banner_rows + 1))->getFont()->setBold(true);
                $sheet->setCellValue($header_arr[$i] . ($banner_rows + 1), $titles[$i]);
            }
            //插数据项
            for ($i = 0; $i < $data_rows; $i++) {
                for ($j = 0; $j < count($datas[$key][$i]); $j++) {
                    $sheet->setCellValueExplicit($header_arr[$j] . ($i + $data_row_start + 1), $datas[$key][$i][$j], DataType::TYPE_STRING);
                }
            }
            //设置行高
            for ($i = 0; $i < ($banner_rows + $data_rows + 1); $i++) {
                $sheet->getRowDimension(($i + 1))->setRowHeight($height);
            }
            //设置居中
            $sheet->getStyle('A1:' . $header_arr[$header_titles_columns - 1] . ($banner_rows + $data_rows + 1))->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            //所有垂直居中
            $sheet->getStyle('A1:' . $header_arr[$header_titles_columns - 1] . ($banner_rows + $data_rows + 1))->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
            //所有自动换行
            $sheet->getStyle('A1:' . $header_arr[$header_titles_columns - 1] . ($banner_rows + $data_rows + 1))->getAlignment()->setWrapText(true);
        }

        //设置excel导出
        $writer = IOFactory::createWriter($php_excel, 'Xls');

        $writer->save($file_path);

        return true;
    }
}
