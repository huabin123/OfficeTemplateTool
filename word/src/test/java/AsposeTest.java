import cn.hutool.core.util.StrUtil;
import com.aspose.words.*;
import com.huabin.common.utils.WordUtils;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Author huabin
 * @DateTime 2022-12-14 14:50
 * @Desc
 */
public class AsposeTest {

    private Document doc;

    @Test
    public void testIfConditionWithMailMerge() throws Exception {

        // 验证License
        if (!WordUtils.getLicense()) {
            return;
        }

        // 模版地址
//        String docxPath = "/Users/huabin/Desktop/template.docx";
//        String docxPath = "/Users/huabin/workspace/playground/工具向/OfficeTemplateTool/word/src/main/resources/doc/template/template-01.docx";
        String docxPath = "/Users/huabin/Downloads/Mail merge destinations - Fax (1).docx";
//        String docxPath = "/Users/huabin/workspace/playground/工具向/OfficeTemplateTool/word/src/main/resources/doc/template/附件：理财产品财务报告附注模板_20221118（更新版）.docx";

        // 读取word模板
        FileInputStream fileInputStream = new FileInputStream(docxPath);
        this.doc = new Document(fileInputStream);

//        String[] fieldNames = {
//                "param1"
//        };
//
//        Object[] fieldValues = {
//                true
//        };

        String[] fieldNames = new String[]{"RecipientName", "SenderName", "FaxNumber", "PhoneNumber",
                "Subject", "Body", "Urgent", "ForReview", "PleaseComment", "param1"};
        Object[] fieldValues = new Object[]{"Josh", "Jenny", "123456789", "", "Hello",
                "<b>HTML Body Test message 1</b>", true, false, true, true};

        // 处理if...else
        doc.getMailMerge().execute(fieldNames, fieldValues);

        // 保存到本地
        doc.save("FinalFile.docx", SaveFormat.DOCX);
    }

    @Test
    public void testIfCondition() throws Exception {
        // 验证License
        if (!WordUtils.getLicense()) {
            return;
        }

        // 模版地址
        String docxPath = "src/main/resources/doc/template/template-01.docx";

        // 读取word模板
        FileInputStream fileInputStream = new FileInputStream(docxPath);
        this.doc = new Document(fileInputStream);

        // 获取单维指标
        Map<String, String> singleData = this.getSingleData();

        //

    }

    @Test
    public void testAspose() throws Exception {

        // 验证License
        if (!WordUtils.getLicense()) {
            return;
        }

        // 模版地址
        String docxPath = "src/main/resources/doc/template/template-01.docx";
//        String docxPath = "/Users/huabin/workspace/playground/工具向/OfficeTemplateTool/word/src/main/resources/doc/template/附件：理财产品财务报告附注模板_20221118（更新版）.docx";

        // 读取word模板
        FileInputStream fileInputStream = new FileInputStream(docxPath);
        this.doc = new Document(fileInputStream);

        // 读取配置表（最后一张表）获取指标配置，读取完成后删除配置表
        NodeCollection allTables = doc.getChildNodes(NodeType.TABLE, true);

        if (allTables.getCount() == 0) {
            throw new Exception("模板有误，找不到表格！");
        }

        Table lastTable = (Table)allTables.get(allTables.getCount() - 1);

        RowCollection tableRows = lastTable.getRows();
        HashMap<String, List> indiListMap = new HashMap<>();
        ArrayList<String> singleIndicatorList = new ArrayList<>();
        ArrayList<String> multiIndicatorList = new ArrayList<>();
        String singleIndicatorPatternStr = "(^SF_VFUN|^.*SF_VFUN)(\\()(\")([^,]+)(\")(,?)(.*)(\\)$)";
        String multiIndicatorPatternStr = "(^SF_MDFUN|^.*SF_MDFUN)(\\()(\")(.*)(\",){1}";
        for (Row tableRow : tableRows) {
            String formula = tableRow.getLastCell().getText().replaceAll("\u0007","");
            String funcCode = "";
            // 单位指标
            if (formula.startsWith("=SF_VFUN")) {
                // 解析公式，获取指标code
                Pattern pattern = Pattern.compile(singleIndicatorPatternStr);
                Matcher matcher = pattern.matcher(formula);

                if (matcher.find()) {
                    funcCode = matcher.group(4);
                }
                singleIndicatorList.add(funcCode);
            } else if (formula.startsWith("=SF_MDFUN")) {
                Pattern pattern = Pattern.compile(multiIndicatorPatternStr);
                Matcher matcher = pattern.matcher(formula);
                if (matcher.find()) {
                    funcCode = matcher.group(4);
                }
                multiIndicatorList.add(funcCode);
            }
        }

        indiListMap.put("single", singleIndicatorList);
        indiListMap.put("multi", multiIndicatorList);



        // 调取外部接口，获取数据
        // 删除最后一张表
        lastTable.remove();

        // 填充段落中的数据
        Map<String, String> singleData = this.getSingleData();
        for (Map.Entry<String, String> entry : singleData.entrySet()) {
            String value = entry.getValue();
            if (value.contains("\n")) {
                value = value.replaceAll("\n", ControlChar.LINE_BREAK);
            }
            doc.getRange().replace(Pattern.compile(entry.getKey()), value);
        }

        // 填充表格中的数据
        Map<String, List<List<String>>> mulData = this.getMulData();
        for (int i = 0; i < allTables.getCount(); i++) {
            Table table = (Table) allTables.get(i);
            RowCollection rows = table.getRows();
                int startRow = 0;
                for (Row row1 : rows) {
                    CellCollection cells = row1.getCells();
                    int startCell = 0;
                    for (Cell cell : cells) {
                        String cellText = StrUtil.trim(cell.getText()).replaceAll("\u0007", "");
                        if (mulData.containsKey(cellText)) {
                            List<List<String>> rowList = mulData.get(cellText);
                            this.fillTableRows(table, startRow, startCell, rowList);
                        } else if (singleData.containsKey(cellText)) {
                            replaceCell(cell, cellText);
                        }
                        startCell++;
                    }
                    startRow++;
                }
        }

        // 处理跨页的情况
//        LayoutCollector collector = new LayoutCollector(doc);
//        for (int i = 0; i < allTables.getCount(); i++) {
//            Table table = (Table) allTables.get(i);
//            // 表格存在跨页
//            if (collector.getEndPageIndex(table.getLastRow()) - collector.getStartPageIndex(table.getFirstRow()) > 1) {
//                splitTable(table, collector);
//                doc.updatePageLayout();
//            }
//        }

        // 保存到本地
        doc.save("FinalFile.docx", SaveFormat.DOCX);
    }

    /**
     * 分割跨页的表格
     * todo 只处理了一次，需要循环处理
     * @see <a href="https://forum.aspose.com/t/how-to-insert-paragraph-before-table-continuation-with-headingformat/246739/25">How to insert paragraph before Table continuation</a>
     * @param table
     * @param collector
     * @throws Exception
     */
    private void splitTable(Table table, LayoutCollector collector) throws Exception {
        int startPageIndex = collector.getStartPageIndex(table.getFirstRow());

        int breakIndex = -1;
        int firstDataRowIndex = -1;
        // Determine index of row where page breaks. And index of the first data row.
        for (int i = 1; i < table.getRows().getCount(); i++) {
            Row r = table.getRows().get(i);
            if (!r.getRowFormat().getHeadingFormat() && firstDataRowIndex < 0)
                firstDataRowIndex = i;

            int rowPageIndex = collector.getEndPageIndex(r);
            if (rowPageIndex > startPageIndex) {
                breakIndex = i;
                break;
            }
        }

        if (breakIndex > 0) {
            Table clone = (Table) table.deepClone(true);

            // Insert a cloned table after the main table.
            Paragraph para = new Paragraph(doc);
            para.getParagraphFormat().setPageBreakBefore(true);
            para.appendChild(new Run(doc, "续表："));

            table.getParentNode().insertAfter(para, table);
            para.getParentNode().insertAfter(clone, para);

            // Remove content after the breaking row from the main table.
            while (table.getRows().getCount() > breakIndex)
                table.getLastRow().remove();

            // Remove rows before the breaking row from the clonned table.
            for (int i = 1; i < breakIndex; i++)
                clone.getRows().removeAt(firstDataRowIndex);

        }
    }

    /**
     * 填充多维指标列表
     * @param table
     * @param startRowIndex 开始填充的行的下标
     * @param startCellIndex 开始填充的列的下标
     * @param rowList 要填充的数据
     * @throws Exception
     */
    private void fillTableRows(Table table, int startRowIndex, int startCellIndex, List<List<String>> rowList) throws Exception {
        Row startRow = table.getRows().get(startRowIndex);
        for (int i = 0; i < rowList.size(); i++) {
            Row row = table.getRows().get(startRowIndex + i);
            if (row == null) {
                // 创建新行
                Row newRow = new Row(doc);
                table.getRows().add(newRow);
                row = table.getLastRow();
                for (int k = 0; k < startCellIndex; k++) {
                    // 补充cell
                    genCell(startRow.getCells().get(k).getText(), row);
                }
            }
            List<String> rowData = rowList.get(i);
            for (int j = 0; j < rowData.size(); j++) {
                // 从startCell开始填充数据
                Cell cell = row.getCells().get(startCellIndex + j);
                String cellVaule = rowList.get(i).get(j);
                if (cell != null) {
                    replaceCell(cell, cellVaule);
                } else {
                    // 创建cell
                    genCell(cellVaule, row);
                }
            }

        }

    }

    /**
     * 删除单元格中原有内容，替换成指标数据
     * todo 设置中文样式
     * @param cell
     * @param cellValue
     * @throws Exception
     */
    private void replaceCell(Cell cell, String cellValue) throws Exception {
        cell.removeAllChildren();
        Paragraph cell1P = new Paragraph(doc);
        Run cellRun = new Run(doc);
        cellRun.setText(cellValue);
        cellRun.getFont().setName("Arial");
        cell1P.appendChild(cellRun);
        cell.appendChild(cell1P);
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

    /**
     * 模拟获取多维指标数据
     * @return
     */
    private Map<String, List<List<String>>> getMulData() {

        HashMap<String, List<List<String>>> multiIndicatorMap = new HashMap<>();

        List<List<String>> multi = new ArrayList<>();
        for (int i = 0; i < 6; i++) {
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

    /**
     * 模拟获取单维指标数据
     * @return
     */
    private Map<String, String> getSingleData(){
        HashMap<String, String> singleMap = new HashMap<>();
        singleMap.put("D0001", "1234");
        singleMap.put("D0011", "5678");
        singleMap.put("D0012", "换行测试\n行2\n行3");
        singleMap.put("param1", "1");
        return singleMap;
    }

}
