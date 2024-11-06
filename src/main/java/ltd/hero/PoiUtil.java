package ltd.hero;


import cn.hutool.core.io.file.FileNameUtil;
import cn.hutool.core.util.StrUtil;
import lombok.SneakyThrows;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;


/**
 * @className: PoiUtil
 * @description: TODO 类描述
 * @author: administrator
 * @date: 2024/11/6
 **/

public class PoiUtil {
    public static void main(String[] args) {
        String filename = "E:\\dev\\java\\replenish-xlsx\\replenish-xlsx\\test.xlsx";
        fixExcel(filename);
    }

    /**
     * 从右向左，从上到下
     *
     * @param filename
     */
    @SneakyThrows
    public static String fixExcel(String filename) {
        FileInputStream input = new FileInputStream(filename);
        Workbook workbook = WorkbookFactory.create(input);
        Sheet sheet = workbook.getSheetAt(0);
        Row firstRow = sheet.getRow(0);
        // 使用 getLastCellNum 获取总列数
        int columnCountByLastCellNum = firstRow.getLastCellNum();
        System.out.println("总列数: " + columnCountByLastCellNum);
        int lastRowNum = sheet.getLastRowNum();
        for (int i = columnCountByLastCellNum - 2; i >= 0; i--) {
            for (int j = 2; j <= lastRowNum; j++) {
                Row row = sheet.getRow(j);
                Cell cell = row.getCell(i);
                Cell cellRight = row.getCell(i + 1);
                if (null != cellRight && StrUtil.isNotEmpty(cellRight.toString())) {
                    Cell cell1 = sheet.getRow(j - 1).getCell(i);
                    cell1.setCellType(CellType.STRING);
                    if (cell == null) {
                        row.createCell(i).setCellValue(cell1.getStringCellValue());
                    } else if (StrUtil.isEmptyIfStr(cell.getStringCellValue())) {
                        cell.setCellValue(cell1.getStringCellValue());
                    }
                }
            }
        }
        String out = new File(filename).getParent() + File.separator + FileNameUtil.mainName(filename) + "-fixed" + ".xlsx";
        FileOutputStream fileOut = new FileOutputStream(out);
        workbook.write(fileOut);
        fileOut.close();
        return out;
    }
}
