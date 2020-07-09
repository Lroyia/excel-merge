package com;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

/**
 *
 * 合并多个相同表头的多个excel文件的程序
 *
 * 运行主类
 * @author <a href="https://blog.lroyia.top">lroyia</a>
 * @since 2020/7/9 15:51
 **/
public class Main {

    /**
     * 扫描的本地路径
     */
    private static final String SCAN_PATH = "C:\\Users\\81487\\Documents\\WeChat Files\\wxid_jzotawfh0qt741\\FileStorage\\File\\2020-07\\20191.1至2020.5.1";

    /**
     * 导出文件名
     */
    private static final String EXPORT_FILE_NAME = "信访内容合并.xls";

    /**
     * 是否有表头
     */
    private static final boolean HAS_HEADER = true;

    public static void main(String[] args) throws Exception {
        // 判断路径
        File path = new File(SCAN_PATH);
        if(!path.isDirectory()) throw new IOException("请填写正确的扫描路径");

        File[] files = path.listFiles((file, name)->!name.equals(EXPORT_FILE_NAME));

        Workbook book = new HSSFWorkbook();

        int rowIndex = 0;
        for (int i = 0; i < files.length; i++) {
            File curFile = files[i];
            System.out.println("读取"+curFile.getName());
            // 从第一个文件确定表头和sheet
            Workbook curBook = WorkbookFactory.create(curFile);
            Iterator<Sheet> iterator = curBook.sheetIterator();
            int index = -1;
            while (iterator.hasNext()){
                Sheet curSheet = iterator.next();
                index++;
                Sheet sheet;
                if(i == 0){
                    sheet = book.createSheet(curSheet.getSheetName());
                    if(HAS_HEADER){
                        Row header = curSheet.getRow(0);
                        Row bookHeader = sheet.createRow(0);
                        bookHeader.setHeightInPoints(header.getHeightInPoints());
                        for(int j = 0; j < header.getLastCellNum(); j++){
                            Cell curCell = header.getCell(j);
                            String value = curCell.getStringCellValue();
                            Cell cell = bookHeader.createCell(j);
                            CellStyle style = cell.getCellStyle();
                            cloneCellStyle(style, curCell.getCellStyle());
                            setCellValue(cell, value, style);
                        }
                        rowIndex = 1;
                    }
                }else{
                    sheet = book.getSheetAt(index);
                }
                for(int j = HAS_HEADER ? 1 : 0; j < curSheet.getLastRowNum(); j++){
                    Row row = sheet.createRow(rowIndex++);
                    Row curRow = curSheet.getRow(j);
                    row.setHeightInPoints(curRow.getHeightInPoints());
                    for(int k = 0; k < curRow.getLastCellNum(); k++){
                        Cell curCell = curRow.getCell(k);
                        String value = curCell.getStringCellValue();
                        Cell cell = row.createCell(k);
                        CellStyle style = cell.getCellStyle();
                        cloneCellStyle(style, curCell.getCellStyle());
                        setCellValue(cell, value, style);
                    }
                }
            }
        }

        File file = new File(SCAN_PATH + File.separator + EXPORT_FILE_NAME);
        if(file.exists()) file.delete();

        FileOutputStream os = new FileOutputStream(file);
        book.write(os);
        os.flush();
        os.close();

    }

    private static void setCellValue(Cell cell, String value, CellStyle cellStyle){
        if(cell == null) return;
        cell.setCellValue(value);
        if(cellStyle != null) cell.setCellStyle(cellStyle);
    }

    private static void cloneCellStyle(CellStyle target, CellStyle source){
        if(target == null || source == null) return;
        target.setAlignment(source.getAlignment());
        target.setBorderBottom(source.getBorderBottom());
        target.setBorderRight(source.getBorderBottom());
        target.setBorderLeft(source.getBorderLeft());
        target.setBorderTop(source.getBorderTop());
        target.setBottomBorderColor(source.getBottomBorderColor());
        target.setFillBackgroundColor(source.getFillBackgroundColor());
        target.setFillForegroundColor(source.getFillBackgroundColor());
        target.setDataFormat(source.getDataFormat());
        target.setFillPattern(source.getFillPattern());
        target.setVerticalAlignment(source.getVerticalAlignment());
    }
}
