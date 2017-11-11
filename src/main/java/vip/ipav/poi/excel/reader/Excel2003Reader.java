package vip.ipav.poi.excel.reader;

/**
 * Created by 89003522 on 2017/11/2.
 */

import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataFormatter;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 抽象Excel2003读取器，通过实现HSSFListener监听器，采用事件驱动模式解析excel2003
 * 中的内容，遇到特定事件才会触发，大大减少了内存的使用。
 */
public class Excel2003Reader implements HSSFListener {
    private int minColumns = -1;
    private POIFSFileSystem fs;
    private int lastRowNumber;
    private int lastColumnNumber;

    /**
     * 是否需要输出值还是输出公式，默认输出值
     */
    private boolean outputFormulaValues = true;

    /**
     * For parsing Formulas
     */
    private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;
    //excel2003工作薄
    private HSSFWorkbook stubWorkbook;

    // Records we pick up as we process
    private SSTRecord sstRecord;
    private FormatTrackingHSSFListener formatListener;

    //表索引
    private int sheetIndex = -1;
    private Integer readSheetIndex = null;
    private BoundSheetRecord[] orderedBSRs;
    private ArrayList boundSheetRecords = new ArrayList();

    //日期格式处理
    private final DataFormatter formatter = new DataFormatter();

    // For handling formulas with string results
    private int nextRow;
    private int nextColumn;
    private boolean outputNextStringRecord;
    //当前行
    private int curRow = 0;
    // 开始读数据的行数
    private int beginRow;
    // 结束读数据的行数
    private int endRow;
    //存储行记录的容器
    private List<Object> rowList = new ArrayList<Object>();

    private String sheetName;

    private List<List<Object>> allValueList = new ArrayList<>();

    public List<List<Object>> getAllValueList() {
        return allValueList;
    }

    public Excel2003Reader(int beginRow, String path) throws IOException {
        this.beginRow = beginRow;
        this.fs = new POIFSFileSystem(new FileInputStream(path));
    }

    public Excel2003Reader(int beginRow, int rows, String path) throws IOException {
        this.beginRow = beginRow;
        this.endRow = this.beginRow + rows - 1;
        this.fs = new POIFSFileSystem(new FileInputStream(path));
    }

    public Excel2003Reader(int beginRow, InputStream in) throws IOException {
        this.beginRow = beginRow;
        this.fs = new POIFSFileSystem(in);
    }

    public Excel2003Reader(int beginRow, int rows, InputStream in) throws IOException {
        this.beginRow = beginRow;
        this.endRow = this.beginRow + rows - 1;
        this.fs = new POIFSFileSystem(in);
    }

    /**
     * 遍历excel下所有的sheet
     *
     * @throws IOException
     */
    public void processAllSheets() throws IOException {
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(
                this);
        formatListener = new FormatTrackingHSSFListener(listener);
        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();
        if (outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(
                    formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }
        factory.processWorkbookEvents(request, fs);
    }

    /**
     * 读取指定表id的数据
     *
     * @param rId rId1,rId2对应sheet1和sheet2
     * @throws Exception
     */
    public void processOneSheet(Integer rId) throws Exception {
        this.readSheetIndex = rId;
        this.processAllSheets();
    }

    /**
     * HSSFListener 监听方法，处理 Record
     */
    public void processRecord(Record record) {
        int thisRow = -1;
        int thisColumn = -1;
        String thisStr = null;
        String value = null;

        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add(record);
                break;

            case BOFRecord.sid:
                BOFRecord br = (BOFRecord) record;
                if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                    // 如果有需要，则建立子工作薄
                    if (workbookBuildingListener != null && stubWorkbook == null) {
                        stubWorkbook = workbookBuildingListener
                                .getStubHSSFWorkbook();
                    }
                    sheetIndex++;
                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord
                                .orderByBofPosition(boundSheetRecords);
                    }
                    sheetName = orderedBSRs[sheetIndex].getSheetname();
                }
                break;

            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;

            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;
                thisRow = brec.getRow();
                thisColumn = brec.getColumn();
                thisStr = "";
                rowList.add(thisColumn, thisStr);
                break;

            case BoolErrRecord.sid: //单元格为布尔类型
                BoolErrRecord berec = (BoolErrRecord) record;
                thisRow = berec.getRow();
                thisColumn = berec.getColumn();
                thisStr = berec.getBooleanValue() + "";
                rowList.add(thisColumn, thisStr);
                break;

            case FormulaRecord.sid: //单元格为公式类型
                FormulaRecord frec = (FormulaRecord) record;
                thisRow = frec.getRow();
                thisColumn = frec.getColumn();
                if (outputFormulaValues) {
                    if (Double.isNaN(frec.getValue())) {
                        // Formula result is a string
                        // This is stored in the next record
                        outputNextStringRecord = true;
                        nextRow = frec.getRow();
                        nextColumn = frec.getColumn();
                    } else {
                        thisStr = formatListener.formatNumberDateCell(frec);
                    }
                } else {
                    thisStr = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook,
                            frec.getParsedExpression()) + '"';
                }
                rowList.add(thisColumn, thisStr);
                break;
            case StringRecord.sid://单元格中公式的字符串
                if (outputNextStringRecord) {
                    // String for formula
                    StringRecord srec = (StringRecord) record;
                    thisStr = srec.getString();
                    thisRow = nextRow;
                    thisColumn = nextColumn;
                    outputNextStringRecord = false;
                }
                break;
            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;
                curRow = thisRow = lrec.getRow();
                thisColumn = lrec.getColumn();
                value = lrec.getValue().trim();
                value = value.equals("") ? null : value;
                this.rowList.add(thisColumn, value);
                break;
            case LabelSSTRecord.sid:  //单元格为字符串类型
                LabelSSTRecord lsrec = (LabelSSTRecord) record;
                curRow = thisRow = lsrec.getRow();
                thisColumn = lsrec.getColumn();
                if (sstRecord == null) {
                    rowList.add(thisColumn, null);
                } else {
                    value = sstRecord
                            .getString(lsrec.getSSTIndex()).toString().trim();
                    value = value.equals("") ? null : value;
                    rowList.add(thisColumn, value);
                }
                break;
            case NumberRecord.sid:  //单元格为数字类型(含日期类型)
                NumberRecord numrec = (NumberRecord) record;
                curRow = thisRow = numrec.getRow();
                thisColumn = numrec.getColumn();
                //HSSFDateUtil.isCellDateFormatted(numrec);
                if("m/d/yy" == formatListener.getFormatString(numrec)){
                    //full format is "yyyy-MM-dd hh:mm:ss.SSS";
                    value = formatter.formatRawCellContents(numrec.getValue(),
                            formatListener.getFormatIndex(numrec), "yyyy-MM-dd");
                }else{
                    value = formatListener.formatNumberDateCell(numrec).trim();
                    value = value.equals("")?"":value;
                }
                // 向容器加入列
                rowList.add(thisColumn, value);
                break;
            default:
                break;
        }

        // 遇到新行的操作
        if (thisRow != -1 && thisRow != lastRowNumber) {
            lastColumnNumber = -1;
        }

        // 空值的操作
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
            curRow = thisRow = mc.getRow();
            thisColumn = mc.getColumn();
            rowList.add(thisColumn, null);
        }

        // 更新行和列的值
        if (thisRow > -1)
            lastRowNumber = thisRow;
        if (thisColumn > -1)
            lastColumnNumber = thisColumn;

        // 行结束时的操作
        if (record instanceof LastCellOfRowDummyRecord) {
            if (minColumns > 0) {
                // 列值重新置空
                if (lastColumnNumber == -1) {
                    lastColumnNumber = 0;
                }
            }
            lastColumnNumber = -1;
            // 每行结束时,执行的数据，临时解决办法
            if (curRow + 1 < beginRow || (endRow > 0 && curRow + 1 > endRow)) {
                //不保存数据，临时解决办法
                rowList.clear();
            } else {
                if (readSheetIndex != null && readSheetIndex > 0) {
                    if (readSheetIndex - 1 == sheetIndex) {
                        this.allValueList.add(rowList);
                        this.rowList = new ArrayList<>();
                    }
                    rowList.clear();
                } else {
                    this.allValueList.add(rowList);
                    this.rowList = new ArrayList<>();
                }
            }
        }
    }
}