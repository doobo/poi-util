package vip.ipav.poi.excel.reader;

/**
 * Created by 89003522 on 2017/11/2.
 */

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * 事件驱动模式处理Excel，针对xlsx文件
 * 只匹配普通非图表类的电子表格
 * 空单元格以null表示
 */
public class Excel2007Reader {

    //获取数据源，可支持path，InputStream
    private OPCPackage pkg = null;

    private static StylesTable stylesTable;
    // 开始读数据的行数
    private int beginRow;
    // 结束读数据的行数
    private int endRow;
    // 所有值列表
    private final static String RID = "rId";
    private List<List<Object>> allValueList = new ArrayList<>();

    //用一个enum表示单元格可能的数据类型
    enum CellDataType {
        BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
    }

    public Excel2007Reader(int beginRow, String path) throws InvalidFormatException {
        this.beginRow = beginRow;
        this.open(path);
    }

    public Excel2007Reader(int beginRow, InputStream in) throws InvalidFormatException, java.io.IOException {
        this.beginRow = beginRow;
        this.open(in);
    }

    public Excel2007Reader(int beginRow, int rows, String path) throws InvalidFormatException {
        this.beginRow = beginRow;
        this.endRow = this.beginRow + rows - 1;
        this.open(path);
    }

    public Excel2007Reader(int beginRow, int rows, InputStream in) throws InvalidFormatException, java.io.IOException {
        this.beginRow = beginRow;
        this.endRow = this.beginRow + rows - 1;
        this.open(in);
    }

    public void open(String path) throws InvalidFormatException {
        this.pkg = OPCPackage.open(path, PackageAccess.READ);
    }

    public void open(InputStream in) throws InvalidFormatException, java.io.IOException {
        this.pkg = OPCPackage.open(in);
    }

    /**
     * 读取指定表id的数据
     *
     * @param rId rId1,rId2对应sheet1和sheet2
     * @throws Exception
     */
    public void processOneSheet(Integer rId) throws Exception {
        InputStream sheet = null;
        rId = rId == null ? 1 : rId;
        if (this.pkg == null) {
            throw new RuntimeException("未设置数据源，请使用open设置对应的数据源");
        }
        try {
            XSSFReader r = new XSSFReader(pkg);
            stylesTable = r.getStylesTable();
            SharedStringsTable sst = r.getSharedStringsTable();
            XMLReader parser = fetchSheetParser(sst);
            sheet = r.getSheet(RID + rId);
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
        } catch (Exception e) {
            throw e;
        } finally {
            if (sheet != null) {
                sheet.close();
            }
        }
    }

    /**
     * 读取文件里面的所有数据
     *
     * @throws Exception
     */
    public void processAllSheets() throws Exception {
        XSSFReader r = new XSSFReader(pkg);
        stylesTable = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = fetchSheetParser(sst);
        Iterator<InputStream> sheets = r.getSheetsData();
        while (sheets.hasNext()) {
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
        }
    }

    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        ContentHandler handler = new SheetHandler(sst);
        parser.setContentHandler(handler);
        return parser;
    }

    public List<List<Object>> getAllValueList() {
        return allValueList;
    }

    private class SheetHandler extends DefaultHandler {
        private SharedStringsTable sst;
        private String lastContents, cs;
        private boolean isString;
        private boolean validRow;
        private int curRow = 0;
        // 定义前一个元素和当前元素的位置，用来计算其中空的单元格数量，如A6和A8等
        private String preRef = null, ref = null;
        // 定义该文档一行最大的单元格数，用来补全一行最后可能缺失的单元格
        private String maxRef = null;

        // 一行的所有数据
        private List<Object> rowValueList;

        private SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            if (name.equals("row") || name.equals("c")) {
                if (preRef == null) {
                    preRef = attributes.getValue("r");
                } else {
                    preRef = ref;
                }
                // 当前单元格的位置
                ref = attributes.getValue("r");
                this.setNextDataType(attributes);

                int column = getColumn(attributes);
                if (column < beginRow || (endRow > 0 && column > endRow)) {
                    validRow = false;
                } else {
                    validRow = true;
                    if (name.equals("row")) {
                        rowValueList = new ArrayList<>();
                        allValueList.add(rowValueList);
                        lastName = null;
                    }
                    String cellType = attributes.getValue("t");
                    if (cellType != null && cellType.equals("s")) {
                        isString = true;
                    } else {
                        isString = false;
                    }
                }
            }
            lastContents = "";
        }

        private String lastName;

        public void endElement(String uri, String localName, String name) throws SAXException {
            if (validRow) {
                if (lastName != null && lastName.equals("c") && name.equals("c")) {
                    rowValueList.add("");
                } else if (isString) {
                    int idx = Integer.parseInt(lastContents);
                    lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                    isString = false;
                    validRow = false;
                    if (name.equals("v")) {//匹配字符串
                        rowValueList.add(this.getDataValue(lastContents.trim(), ""));
                    }
                } else {
                    if (name.equals("c")) {//匹配非字符串(数字、空)或合并的非空单元格
                        rowValueList.add(this.getDataValue(lastContents.trim(), ""));
                    }
                }

                // 补全单元格之间的空单元格
                if ("v".equals(name) || "t".equals(name)) {
                    if (!ref.equals(preRef)) {
                        int len = countNullCell(ref, preRef);
                        for (int i = 0; i < len; i++) {
                            rowValueList.add(null);
                        }
                    }
                } else {
                    // 如果标签名称为 row，这说明已到行尾，调用 optRows() 方法
                    if (name.equals("row")) {
                        String value = "";
                        // 默认第一行为表头，以该行单元格数目为最大数目
                        if (curRow == 0) {
                            maxRef = ref;
                        }
                        // 补全一行尾部可能缺失的单元格
                        if (maxRef != null) {
                            int len = countNullCell(maxRef, ref);
                            for (int i = 0; i <= len; i++) {
                                rowValueList.add(null);
                            }
                        }
                        curRow++;
                        preRef = null;
                        ref = null;
                    }
                }
            }
            lastName = name;
        }

        private CellDataType nextDataType = CellDataType.SSTINDEX;
        private final DataFormatter formatter = new DataFormatter();
        private short formatIndex;
        private String formatString;

        /**
         * 根据element属性设置数据类型
         *
         * @param attributes
         */
        public void setNextDataType(Attributes attributes) {

            nextDataType = CellDataType.NUMBER;
            formatIndex = -1;
            formatString = null;
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if ("b".equals(cellType)) {
                nextDataType = CellDataType.BOOL;
            } else if ("e".equals(cellType)) {
                nextDataType = CellDataType.ERROR;
            } else if ("inlineStr".equals(cellType)) {
                nextDataType = CellDataType.INLINESTR;
            } else if ("s".equals(cellType)) {
                nextDataType = CellDataType.SSTINDEX;
            } else if ("str".equals(cellType)) {
                nextDataType = CellDataType.FORMULA;
            }
            if (cellStyleStr != null) {
                int styleIndex = Integer.parseInt(cellStyleStr);
                XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                formatIndex = style.getDataFormat();
                formatString = style.getDataFormatString();
                if ("m/d/yy" == formatString) {
                    nextDataType = CellDataType.DATE;
                    //full format is "yyyy-MM-dd hh:mm:ss.SSS";
                    formatString = "m/d/yy";
                }
                if (formatString == null) {
                    nextDataType = CellDataType.NULL;
                    formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
                }
            }
        }

        /**
         * 根据数据类型获取数据
         *
         * @param value
         * @param thisStr
         * @return
         */
        public String getDataValue(String value, String thisStr) {
            if(value == null || value.isEmpty()){
                return "";
            }
            switch (nextDataType) {
                //这几个的顺序不能随便交换，交换了很可能会导致数据错误
                case BOOL:
                    char first = value.charAt(0);
                    thisStr = first == '0' ? "FALSE" : "TRUE";
                    break;
                case ERROR:
                    thisStr = "\"ERROR:" + value.toString() + '"';
                    break;
                case FORMULA:
                    thisStr = '"' + value.toString() + '"';
                    break;
                case INLINESTR:
                    XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                    thisStr = rtsi.toString();
                    rtsi = null;
                    break;
                case SSTINDEX:
                    String sstIndex = value.toString();
                    thisStr = value.toString();
                    break;
                case NUMBER:
                    if (formatString != null) {
                        thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
                    } else {
                        thisStr = value;
                    }
                    thisStr = thisStr.replace("_", "").trim();
                    break;
                case DATE:
                    try {
                        thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
                    } catch (NumberFormatException ex) {
                        thisStr = value.toString();
                    }
                    thisStr = thisStr.replace(" ", "");
                    break;
                default:
                    thisStr = "";
                    break;
            }
            return thisStr;
        }

        /**
         * 计算两个单元格之间的单元格数目(同一行)
         *
         * @param ref
         * @param preRef
         * @return
         */
        public int countNullCell(String ref, String preRef) {
            // excel2007最大行数是1048576，最大列数是16384，最后一列列名是XFD
            String xfd = ref.replaceAll("\\d+", "");
            String xfd_1 = preRef.replaceAll("\\d+", "");

            xfd = fillChar(xfd, 3, '@', true);
            xfd_1 = fillChar(xfd_1, 3, '@', true);

            char[] letter = xfd.toCharArray();
            char[] letter_1 = xfd_1.toCharArray();
            int res = (letter[0] - letter_1[0]) * 26 * 26
                    + (letter[1] - letter_1[1]) * 26
                    + (letter[2] - letter_1[2]);
            return res - 1;
        }

        /**
         * 字符串的填充
         *
         * @param str
         * @param len
         * @param let
         * @param isPre
         * @return
         */
        String fillChar(String str, int len, char let, boolean isPre) {
            int len_1 = str.length();
            if (len_1 < len) {
                if (isPre) {
                    for (int i = 0; i < (len - len_1); i++) {
                        str = let + str;
                    }
                } else {
                    for (int i = 0; i < (len - len_1); i++) {
                        str = str + let;
                    }
                }
            }
            return str;
        }

        public void characters(char[] ch, int start, int length) throws SAXException {
            lastContents += new String(ch, start, length);
        }

        private int getColumn(Attributes attributes) {
            String row = attributes.getValue("r");
            int firstDigit = -1;
            for (int c = 0; c < row.length(); ++c) {
                if (Character.isDigit(row.charAt(c))) {
                    firstDigit = c;
                    break;
                }
            }
            return Integer.valueOf(row.substring(firstDigit, row.length()));
        }
    }
}
