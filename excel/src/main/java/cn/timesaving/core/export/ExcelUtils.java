/**
 * Mainbo.com Inc.
 * Copyright (c) 2015-2017 All Rights Reserved.
 */

package cn.timesaving.core.export;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.UUID;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * <pre>
 * Excle操作的工具
 * </pre>
 *
 * @author wangyao,tmser
 * @version $Id: ExcleUtils.java, v 1.0 2015年10月16日 下午5:15:40 wangyao,tmser Exp
 *          $
 */
public abstract class ExcelUtils {

  public static final String FORMART_DATE = "yyyy-mm-dd";
  public static final String END_TIME = "2999-1-2";
  public static final String START_TIME = "1900-1-2";

  public static final String DEFAULT_NAME_PRE = "tmpt_row_";// 数据字典sheet

  public static final String DEFAULT_HIDDEN_SHEET = "hidden";// 隐藏sheet

  public static enum ExcelType {
    EXL_2003,
    EXL_2007;

    /**
     * parse excel type
     * 
     * @param filename
     *          file name
     * @return excel type
     */
    public static ExcelType parseExcelType(String filename) {
      if (filename != null && filename.trim().toLowerCase().lastIndexOf(".xls") == filename.length() - 4) {
        return ExcelType.EXL_2003;
      } else {
        return ExcelType.EXL_2007;
      }
    }
  }

  /**
   * 无效的名称
   */
  private static final Pattern NAME_VALID = Pattern.compile("^[cr]$|\\s+|^[\\d]|[^\u4e00-\u9fa5\\w_\\.\\\\]");

  /**
   * create work book from file path
   * 
   * @param filePath
   *          file path
   * @return workbook.
   */
  public static Workbook createWorkBook(String filePath) {
    try {
      String ext = filePath.substring(filePath.lastIndexOf(".") + 1);
      if ("xls".equals(ext)) { // 使用xls方式读取
        return new HSSFWorkbook(new FileInputStream(filePath));
      } else if ("xlsx".equals(ext)) { // 使用xlsx方式读取
        return new XSSFWorkbook(new FileInputStream(filePath));
      }
    } catch (Exception e) {
      // logger.error("open {} failed!", filePath);
      // logger.error("", e);
    }
    return null;
  }

  /**
   * create work book from inputstream
   * 
   * @param is
   *          is
   * @param xlsType
   *          excelType
   * @return work book
   */
  public static Workbook createWorkBook(InputStream is, ExcelType xlsType) throws IOException {
    Workbook workbook = null;
    if (xlsType == ExcelType.EXL_2007) {
      workbook = new XSSFWorkbook(is);
    } else {
      workbook = new HSSFWorkbook(is);
    }

    return workbook;
  }

  /**
   * 获取poi excel 中cell 内容
   * 
   * @param cell
   *          cell
   * @return cell value
   */
  public static String getCellValue(Cell cell) {
    if (cell == null) {
      return "";
    }
    String cellValue = "";
    DecimalFormat df = new DecimalFormat("############.##########");
    switch (cell.getCellType()) {
    case Cell.CELL_TYPE_STRING:
      cellValue = cell.getRichStringCellValue().getString();
      break;
    case Cell.CELL_TYPE_NUMERIC:
      if (HSSFDateUtil.isCellDateFormatted(cell)) {
        double value = cell.getNumericCellValue();
        Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        cellValue = dateFormat.format(date);
      } else {
        cellValue = df.format(cell.getNumericCellValue()).toString();
      }
      break;
    case Cell.CELL_TYPE_BOOLEAN:
      cellValue = String.valueOf(cell.getBooleanCellValue());
      break;
    case Cell.CELL_TYPE_FORMULA:
      cellValue = cell.getCellFormula();
      break;
    default:
      cellValue = "";
    }
    return cellValue == null ? "" : cellValue.trim();
  }

  /**
   * 用于将excel表格中列索引转成列号字母，从A对应1开始
   * 
   * @param index
   *          列索引
   * @return 列号
   */
  public static String indexToColumn(int index) {
    if (index == 0) {
      return "A";
    }
    if (index <= 0) {
      try {
        throw new Exception("Invalid parameter");
      } catch (Exception e) {
        e.printStackTrace();
      }
    }
    index--;
    String column = "";
    do {
      if (column.length() > 0) {
        index--;
      }
      column = ((char) (index % 26 + 'A')) + column;
      index = (index - index % 26) / 26;
    } while (index > 0);
    return column;
  }

  /**
   * 1）不能以数字和问号“？”开头。 2）不能以R、r、C、c单个字母作为名称，R和C在R1C1样式中表示行、列
   * 3）不能包含空格，可以用下划线或点号代替。 4）不能使用除：下划线“_”、点号“.”和反斜杠“\”外的其他符号。问号也允许但不能作为名称开头。
   * 5）字符个数不能超过255。 6）不区分大小写 创建名称
   * 
   * @param wb
   *          work book
   * @param name
   *          name
   * @param expression
   *          expression
   * @return Name
   */
  public static Name createName(Workbook wb, String name, String expression) {
    if (wb == null || !validName(name)) {
      return null;
    }
    Name refer = wb.createName();
    refer.setNameName(name);
    refer.setRefersToFormula(expression);
    return refer;
  }

  /**
   * 校验数据字符串的长度是否大于255
   * 
   * @param textList
   *          text list
   * @return true if big than 255
   */
  public static boolean validateLength(String[] textList) {
    StringBuilder sb = new StringBuilder(textList.length * 16);
    for (int i = 0; i < textList.length; i++) {
      if (i > 0) {
        sb.append('\0'); // list delimiter is the nul char
      }
      sb.append(textList[i]);
    }
    return sb.length() >= 255;
  }

  /**
   * 设置数据有效性（通过列表设置） 数据长度超限制，将自动转换为name 方式
   * 
   * @param dataList
   *          data
   * @param firstRow
   *          first row
   * @param endRow
   *          end row
   * @param firstCol
   *          first column
   * @param endCol
   *          endcolumn
   * @return DataValidation
   */
  public static DataValidation createDataValidation(Sheet sheet, String[] dataList, int firstRow, int endRow,
      int firstCol, int endCol) {
    if (sheet == null) {
      return null;
    }
    DataValidationConstraint constraint = null;
    Workbook workbook = sheet.getWorkbook();
    boolean isOldVersion = true;
    if (sheet instanceof XSSFSheet) {
      isOldVersion = false;
    }
    if (validateLength(dataList)) {
      // 设置下拉列表的内容
      Sheet hidden = workbook.createSheet(DEFAULT_NAME_PRE + UUID.randomUUID());
      ExcelCascadeRecord record = createCascadeRecord(hidden, Arrays.asList(dataList));
      constraint = isOldVersion ? DVConstraint.createFormulaListConstraint(record.getName())
          : new XSSFDataValidationConstraint(ValidationType.LIST, record.getName());
      workbook.setSheetHidden(workbook.getSheetIndex(hidden), true);
    } else {
      constraint = isOldVersion ? DVConstraint.createExplicitListConstraint(dataList)
          : new XSSFDataValidationConstraint(dataList);
      ;
    }

    return createDataValidation(sheet, constraint, firstRow, endRow, firstCol, endCol);
  }

  /**
   * 设置日期单元格数据校验
   * 
   * @param sheet
   * @param firstRow
   *          起始行
   * @param lastRow
   *          终止行
   * @param firstCol
   *          起始列
   * @param lastCol
   *          终止列
   * @param errorTitle
   *          错误提示标题
   * @param errorText
   *          错误提示信息
   * @throws IllegalArgumentException
   */
  public static void setValidationDate(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol,
      String errorTitle, String errorText) throws IllegalArgumentException {
    if (firstRow < 0 || lastRow < 0 || firstCol < 0 || lastCol < 0 || lastRow < firstRow || lastCol < firstCol) {
      throw new IllegalArgumentException(
          "Wrong Row or Column index : " + firstRow + ":" + lastRow + ":" + firstCol + ":" + lastCol);
    }
    if (sheet instanceof XSSFSheet) {
      XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
      DataValidationConstraint dvConstraint = dvHelper.createDateConstraint(OperatorType.BETWEEN, START_TIME, END_TIME,
          FORMART_DATE);
      CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
      XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, addressList);
      validation.setSuppressDropDownArrow(true);
      validation.setShowErrorBox(true);
      validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
      validation.createErrorBox(errorTitle, errorText);
      sheet.addValidationData(validation);
    } else if (sheet instanceof HSSFSheet) {
      CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
      DVConstraint dvConstraint = DVConstraint.createDateConstraint(OperatorType.BETWEEN, START_TIME, END_TIME,
          FORMART_DATE);
      DataValidation validation = new HSSFDataValidation(addressList, dvConstraint);
      validation.setSuppressDropDownArrow(true);
      validation.setShowErrorBox(true);
      validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
      validation.createErrorBox(errorTitle, errorText);
      sheet.addValidationData(validation);
    }
  }

  /**
   * 设置数据有效性（通过列表设置） 数据长度超限制，将自动转换为name 方式
   * 
   * @param sheet
   *          sheet
   * @param name
   *          name
   * @param firstRow
   *          first row
   * @param endRow
   *          end row
   * @param firstCol
   *          first column
   * @param endCol
   *          end column
   * @return DataValidation
   */
  public static DataValidation createDataValidation(Sheet sheet, String name, int firstRow, int endRow, int firstCol,
      int endCol) {
    if (sheet == null) {
      return null;
    }
    DataValidationConstraint constraint = null;
    boolean isOldVersion = true;
    if (sheet instanceof XSSFSheet) {
      isOldVersion = false;
    }
    // 设置下拉列表的内容
    constraint = isOldVersion ? DVConstraint.createFormulaListConstraint(name)
        : new XSSFDataValidationConstraint(ValidationType.LIST, name);
    return createDataValidation(sheet, constraint, firstRow, endRow, firstCol, endCol);
  }

  /**
   * 设置数据有效性（通过@see DVConstraint 设置）
   * 
   * @param sheet
   *          sheet
   * @param constraint
   *          data
   * @param firstRow
   *          first row
   * @param endRow
   *          end row
   * @param firstCol
   *          first column
   * @param endCol
   *          endcolumn
   * @return DataValidation
   */
  public static DataValidation createDataValidation(Sheet sheet, DataValidationConstraint constraint, int firstRow,
      int endRow, int firstCol, int endCol) {
    boolean isOldVersion = true;
    if (sheet instanceof XSSFSheet) {
      isOldVersion = false;
    }
    // logger.debug("起始行:" + firstRow + " - 起始列:" + firstCol + ", 终止行:" + endRow
    // + " - 终止列:" + endCol);
    if (isOldVersion) {
      // 设置下拉列表的内容
      // 四个参数分别是：起始行、终止行、起始列、终止列
      CellRangeAddressList regions = new CellRangeAddressList((short) firstRow, (short) endRow, (short) firstCol,
          (short) endCol);
      // 数据有效性对象
      return new HSSFDataValidation(regions, constraint);
    } else {
      XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
      CellRangeAddressList addressList = new CellRangeAddressList((short) firstRow, (short) endRow, (short) firstCol,
          (short) endCol);
      XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(constraint, addressList);
      validation.setSuppressDropDownArrow(true);
      return validation;
    }
  }

  private static boolean validName(String name) {
    boolean rs = false;
    if (StringUtils.isNotBlank(name)) {
      if (!NAME_VALID.matcher(name.trim()).find()) {
        rs = true;
      }
    }
    return rs;
  }

  /**
   * 创建级联关系
   * 
   * @param workbook
   *          work book
   * @param dataMap
   *          data
   * @return ExcelCascadeRecord
   */
  public static ExcelCascadeRecord createCascadeRecord(Workbook workbook, Map<String, Object> dataMap) {
    return createCascadeRecord(workbook, null, dataMap);
  }

  /**
   * 创建级联关系
   * 
   * @param workbook
   *          work book
   * @param name
   *          name
   * @param dataMap
   *          data
   * @return ExcelCascadeRecord
   */
  public static ExcelCascadeRecord createCascadeRecord(Workbook workbook, String name, Map<String, Object> dataMap) {
    if (workbook == null || dataMap == null) {
      return null;
    }

    Sheet hidden = workbook.createSheet(DEFAULT_HIDDEN_SHEET + UUID.randomUUID());
    ExcelCascadeRecord record = createCascadeRecord(hidden, name, dataMap);
    workbook.setSheetHidden(workbook.getSheetIndex(hidden), true);
    return record;
  }

  /**
   * 使用已有sheet 创建级联关系
   * 
   * @param sheet
   *          sheet
   * @param name
   *          name
   * @param isAppendName
   *          是否对超过两级级联名称进行拼接,如果使用拼接，引用时，也可拼接两列，如INDIRECT($D$18&$E$18)
   * @param dataMap
   *          data
   * @return ExcelCascadeRecord
   */
  public static ExcelCascadeRecord createCascadeRecord(Sheet sheet, String name, boolean isAppendName,
      Map<String, Object> dataMap) {
    if (sheet == null || dataMap == null) {
      return null;
    }

    ExcelCascadeRecord record = createCascadeRecord(sheet, name, dataMap.keySet());
    if (record != null) {
      for (String key : dataMap.keySet()) {
        record.putChildRecord(name, createCascadeRecord(sheet, key, dataMap.get(key), isAppendName));
      }
    }
    return record;
  }

  /**
   * 使用已有sheet 创建级联关系
   * 
   * @param sheet
   *          sheet
   * @param name
   *          name
   * @param dataMap
   *          data
   * @return ExcelCascadeRecord
   */
  public static ExcelCascadeRecord createCascadeRecord(Sheet sheet, String name, Map<String, Object> dataMap) {
    return createCascadeRecord(sheet, name, dataMap, false);
  }

  @SuppressWarnings({ "unchecked", "rawtypes" })
  private static ExcelCascadeRecord createCascadeRecord(Sheet hidden, String name, Object data, boolean isAppendName) {
    ExcelCascadeRecord record = null;
    if (data instanceof Collection) {
      record = createCascadeRecord(hidden, name, (Collection) data);
    } else if (data instanceof Map) {
      Map<String, Object> mapdata = (Map<String, Object>) data;
      record = createCascadeRecord(hidden, name, mapdata.keySet());
      if (record == null) {
        return null;
      }
      String childName;
      for (String key : mapdata.keySet()) {
        if (isAppendName) {
          childName = name + key;
        } else {
          childName = key;
        }
        ExcelCascadeRecord tmpRecord = createCascadeRecord(hidden, childName, mapdata.get(key), isAppendName);
        if (tmpRecord != null) {
          record.putChildRecord(name, tmpRecord);
        }
      }
    }

    return record;
  }

  /**
   * 
   * @param dicSheet
   *          sheet
   * @param datas
   *          data
   * @return ExcelCascadeRecord
   */
  public static ExcelCascadeRecord createCascadeRecord(Sheet dicSheet, Collection<String> datas) {
    return createCascadeRecord(dicSheet, 0, null, datas);
  }

  /**
   * 
   * @param dicSheet
   *          sheet
   * @param name
   *          name
   * @param datas
   *          data
   * @return ExcelCascadeRecord
   */
  public static ExcelCascadeRecord createCascadeRecord(Sheet dicSheet, String name, Collection<String> datas) {
    return createCascadeRecord(dicSheet, 0, name, datas);
  }

  /**
   * 创建级联excel 记录
   * 
   * @param dicSheet
   *          sheet
   * @param startColumn
   *          start column
   * @param name
   *          name
   * @param datas
   *          data
   * @return ExcelCascadeRecord
   */
  public static ExcelCascadeRecord createCascadeRecord(Sheet dicSheet, int startColumn, final String name,
      Collection<String> datas) {
    if (dicSheet == null || datas == null || datas.size() == 0) {
      return null;
    }

    if (startColumn < 0) {
      startColumn = 0;
    }

    String cname = name;
    if (StringUtils.isBlank(name)) {
      cname = DEFAULT_NAME_PRE + UUID.randomUUID();
    }

    int curCellColumn = startColumn;
    int endCellColumn = curCellColumn;
    Row valueRow = getAvailableRow(dicSheet);
    Set<String> dataSet = new HashSet<String>();
    dataSet.addAll(datas);
    int rowIndex = valueRow.getRowNum() + 1;
    int endRowIndex = rowIndex;
    ExcelCascadeRecord record = new ExcelCascadeRecord(cname);

    for (String value : datas) {
      if (StringUtils.isNotBlank(value) && dataSet.contains(value)) {
        valueRow.createCell(curCellColumn);
        valueRow.getCell(curCellColumn).setCellValue(value);
        record.addCellData(new ExcelCascadeRecord.CellData(rowIndex, curCellColumn, value));

        dataSet.remove(value);
        curCellColumn++;
        if (curCellColumn == 255) {
          endCellColumn = 255;
          curCellColumn = startColumn;
          valueRow = getAvailableRow(dicSheet);
          endRowIndex = valueRow.getRowNum() + 1;
        } else {
          endCellColumn = curCellColumn;
        }
      }
    }

    String endColumnName = ExcelUtils.indexToColumn(endCellColumn);// 找到对应列对应的列名
    String startColumnName = ExcelUtils.indexToColumn(startColumn);// 找到对应列对应的列名
    String expression = dicSheet.getSheetName() + "!$" + startColumnName + "$" + rowIndex + ":$" + endColumnName + "$"
        + endRowIndex;
    ExcelUtils.createName(dicSheet.getWorkbook(), cname, expression);
    record.setExpression(expression);

    return record;
  }

  private static Row getAvailableRow(Sheet sheet) {
    Row row = null;
    int begin = sheet.getFirstRowNum();
    int end = sheet.getLastRowNum();
    for (int i = begin; i <= end; i++) {
      if (null == sheet.getRow(i)) {
        row = sheet.createRow(i);
      }
    }

    if (row == null) {
      row = sheet.createRow(end + 1);
    }

    return row;
  }
}
