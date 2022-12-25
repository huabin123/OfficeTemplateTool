import cn.hutool.core.util.StrUtil;
import com.aspose.words.*;
import com.huabin.common.utils.WordUtils;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * @Author huabin
 * @DateTime 2022-12-14 14:50
 * @Desc
 */
public class AsposeTest {

    private Document doc;

    @Test
    public void testAspose() throws Exception {

        // 验证License
        if (!WordUtils.getLicense()) {
            return;
        }

        // 模版地址
        String docxPath = "src/main/resources/doc/template/template-01.docx`";
//        String docxPath = "/Users/huabin/workspace/playground/中银/wordTemplate/doc/finance/附件：理财产品财务报告附注模板_20221118（更新版）.docx";

        // 读取word模板
        FileInputStream fileInputStream = new FileInputStream(docxPath);
        this.doc = new Document(fileInputStream);

        int pageCount = doc.getPageCount();
        for (int i = 0; i < pageCount; i++) {
            PageInfo pageInfo = doc.getPageInfo(i);
        }
        doc.getRange().replace(Pattern.compile("\\$\\{123\\}"), "ziwen中文");

        LayoutCollector collector = new LayoutCollector(doc);

        Table table = doc.getSections().get(0).getBody().getTables().get(0);

//        table.removeChild(table.getLastRow());
//
//        Row row = new Row(doc);
//        table.getRows().add(row);
//
//        Paragraph cell1P = new Paragraph(doc);
//        Cell cell1 = new Cell(doc);
//        Run cellRun = new Run(doc);
//        cellRun.setText("wei");
//        cell1P.appendChild(cellRun);
//        cell1.appendChild(cell1P);
//        row.appendChild(cell1);
//
        if (collector.getStartPageIndex(table.getFirstRow()) != collector.getStartPageIndex(table.getLastRow())) {
            System.out.println(collector.getStartPageIndex(table.getFirstRow()));
            System.out.println(collector.getStartPageIndex(table.getLastRow()));
        }

        Map<String, List<List<String>>> mulData = this.getMulData();
        SectionCollection sections = doc.getSections();
        for (int i = 0; i < sections.getCount(); i++) {
            TableCollection tables = sections.get(i).getBody().getTables();
            for (Table table1 : tables) {
                RowCollection rows = table1.getRows();
                int startRow = 0;
                for (Row row1 : rows) {
                    CellCollection cells = row1.getCells();
                    int startCell = 0;
                    for (Cell cell : cells) {
                        String cellText = StrUtil.trim(cell.getText()).replaceAll("\u0007","");
                        if (mulData.containsKey(cellText)) {
                            List<List<String>> rowList = mulData.get(cellText);
                            this.fillTableRows(table1, startRow, startCell, rowList);
                        }
                        startCell++;
                    }
                    startRow++;
                }
            }
        }
//        TableCollection tables = doc.getSections().get(0).getBody().getTables();
//        for (Table table1 : tables) {
//            RowCollection rows = table1.getRows();
//            int startRow = 0;
//            for (Row row1 : rows) {
//                CellCollection cells = row1.getCells();
//                int startCell = 0;
//                for (Cell cell : cells) {
//                    String cellText = StrUtil.trim(cell.getText()).replaceAll("\u0007","");
//                    if (mulData.containsKey(cellText)) {
//                        List<List<String>> rowList = mulData.get(cellText);
//                        this.fillTableRows(table1, startRow, startCell, rowList);
//                    }
//                    startCell++;
//                }
//                startRow++;
//            }
//        }


        // 保存到本地
        doc.save("FinalFile.docx", SaveFormat.DOCX);

    }

    private void fillTableRows(Table table1, int startRow, int startCell, List<List<String>> rowList) throws Exception {

        DocumentBuilder builder = new DocumentBuilder(doc);

        Row row2 = table1.getRows().get(startRow);
        Cell cell2 = row2.getCells().get(startCell);
        CellFormat cellFormat = cell2.getCellFormat();
        // todo 多维指标为空，把这一行全部置空
        for (int i = 0; i < rowList.size(); i++) {
            Row row = table1.getRows().get(startRow+i);
            if (row == null) {
                // 创建新行
                Row row1 = new Row(doc);
                table1.getRows().add(row1);
                row = table1.getLastRow();
                for (int k = 0; k < startCell; k++) {
                    // 补充cell
                    genCell(row2.getCells().get(k).getText(), row);
                }
            }
            List<String> rowData = rowList.get(i);
            for (int j = 0; j < rowData.size(); j++) {
                // 从startCell开始填充数据
                Cell cell = row.getCells().get(startCell+j);
                if (cell != null) {
                    cell.removeAllChildren();
                    Paragraph cell1P = new Paragraph(doc);
                    Run cellRun = new Run(doc);
                    cellRun.setText(rowList.get(i).get(j));

                    cellRun.getFont().setName("Arial");
                    cellRun.getFont().setName("黑体");
                    cell1P.appendChild(cellRun);
                    cell.appendChild(cell1P);
                } else {
                    // 创建cell
                    genCell(rowList.get(i).get(j), row);
                }
//                Paragraph cell1P = new Paragraph(doc);
//                Cell cell1 = new Cell(doc);
//                Run cellRun = new Run(doc);
//                cellRun.setText(rowList.get(i).get(j));
//                cell1P.appendChild(cellRun);
//                cell1.appendChild(cell1P);
//                row.appendChild(cell1);
            }

        }

    }

    private void genCell(String row2, Row row) throws Exception {
        Paragraph cell1P = new Paragraph(doc);
        Cell cell1 = new Cell(doc);
        Run cellRun = new Run(doc);
        cellRun.setText(row2);
        cellRun.getFont().setName("Arial");
        cell1P.appendChild(cellRun);
        cell1.appendChild(cell1P);
        row.appendChild(cell1);
    }


    private Map<String, List<List<String>>> getMulData(){

        HashMap<String, List<List<String>>> multiIndicatorMap = new HashMap<>();

        List<List<String>> multi = new ArrayList<>();
        for (int i = 0; i < 8; i++) {
            ArrayList<String> list = new ArrayList<>();
            list.add("column列1");
            list.add("column列2");
            list.add("column列3");
            multi.add(list);
            multi.add(list);
            multi.add(list);
            multi.add(list);
            multi.add(list);
        }
        multiIndicatorMap.put("D0002", multi);
        return multiIndicatorMap;
    }

}
