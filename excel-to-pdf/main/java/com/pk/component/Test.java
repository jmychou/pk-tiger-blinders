package com.pk.component;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import be.quodlibet.boxable.BaseTable;
import be.quodlibet.boxable.Row;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.poi.ss.usermodel.Cell;


import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @ClassName Test
 * @Description TODO
 * @Author jmy
 * @Date 2022/7/31 08:58
 */
public class Test {
    public static void main(String[] args) throws IOException {
        //1.读取Excel
        XSSFWorkbook workbook = new XSSFWorkbook("e1.xlsx");
        XSSFSheet sheet = workbook.getSheetAt(0);
        Map<String, List<String>> content = new HashMap<>();
        int i = 0;
        for (org.apache.poi.ss.usermodel.Row row : sheet) {
            List<String> list = new ArrayList<>();
            for (Cell cell : row) {
                list.add(cell.getStringCellValue());
                System.out.print(cell.getStringCellValue() + "   ");
            }
            content.put(String.valueOf(i), list);
            i++;
            System.out.println();
        }
        workbook.close();


        //2. 创建PDF对象
        PDDocument pdDocument = new PDDocument();
        PDPage pdPage = new PDPage(PDRectangle.A4);


        //3.创建表格画板
        float margin = 50;

        float height = PDRectangle.A4.getHeight() - 220;
        float width = PDRectangle.A4.getWidth() - 150;

        float yStartNewPage = pdPage.getMediaBox().getHeight() - (2 * margin);
        float tableWidth = pdPage.getMediaBox().getWidth() - (2 * margin);

        //BaseTable baseTable = new BaseTable(height, height, 0, width, 0, pdDocument, pdPage, true, true);
        BaseTable baseTable = new BaseTable(PDRectangle.A4.getHeight() - 50, PDRectangle.A4.getHeight(), 0, PDRectangle.A4.getWidth(), 0, pdDocument, pdPage, true, true);

        content.forEach((k, v) -> {
            Row<PDPage> row1 = baseTable.createRow(20f);
            v.forEach(val -> row1.createCell(30, val));
        });

        baseTable.draw();


        //4. 写入PDF
        pdDocument.addPage(pdPage);
        pdDocument.save("test.pdf");
        pdDocument.close();

    }
}
