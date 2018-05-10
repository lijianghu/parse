/**
 * Mainbo.com Inc.
 * Copyright (c) 2015-2017 All Rights Reserved.
 */

package cn.timesaving.core.export;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import cn.timesaving.core.export.ExcelUtils.ExcelType;

/**
 * <pre>
 *  Excel 批量导入导出，基础类
 * </pre>
 *
 * @author tmser
 * @version $Id: BatchImportService.java, v 1.0 2016年8月30日 下午4:42:06 tmser Exp $
 */
public abstract class ExcelBatchService implements BatchService {

  private static Logger logger = LoggerFactory.getLogger(TableModel.class);

  private static final String REQUIRED_HEADERS = "_REQUIRED_HEADER";

  public static final String TITLE_PARAMS = "_TITLE_PARAMS_";

  /**
   * 批量导出到excel
   * 
   * @param os
   *          导出到的输出流
   * @param template
   *          excel模板文件，可以为空。
   * @param params
   *          ，导入需要的附加参数
   */

  @Override
  public void exportData(OutputStream os, Map<String, Object> params, final StringBuilder resultMsg, File template) {
    Workbook workbook = null;
    InputStream is = null;
    boolean usetemp = false;
    try {
      if (template != null && template.exists()) {
        is = new FileInputStream(template);
        workbook = ExcelUtils.createWorkBook(is,
            ExcelType.parseExcelType(template.getName()));
        usetemp = true;
      } else {
        workbook = new HSSFWorkbook();
      }

      String name = sheetname();
      Object titles = params.get(TITLE_PARAMS);
      if (titles == null) {
        params.put(TITLE_PARAMS, new HashSet<ExcelTitle>());
      }
      Sheet sheet;
      if (!usetemp) { // 生成全新excel 导出
        if (StringUtils.isBlank(name)) {
          name = "Sheet1";
        }
        sheet = workbook.createSheet(name);
        fillData(sheet, params, resultMsg);
      } else { // 使用模板导出
        if (StringUtils.isNotBlank(name)) {
          sheet = workbook.getSheet(name);
          fillData(sheet, params, resultMsg);
        } else {
          for (int i = 0; i < workbook.getNumberOfSheets(); i++) { // 获取每个Sheet表
            sheet = workbook.getSheetAt(i);
            logger.debug("fill sheet data {} start!", i);
            fillData(sheet, params, resultMsg);
            logger.debug("fill sheet data {} end!", i);
          }
        }
      }

      workbook.write(os);
    } catch (IOException e) {
      logger.error("read excel file failed", e);
    } finally {
      try {
        if (is != null) {
          is.close();
        }
      } catch (IOException e) {
        logger.error("close workbook failed", e);
      }
    }

  }

  protected void fillData(Sheet sheet, Map<String, Object> params, StringBuilder resultMsg) {
    beforeSheetFill(sheet, params, resultMsg);
    fillSheetData(sheet, params, resultMsg);
    endSheetFill(sheet, params, resultMsg);
  }

  /**
   * 填充sheet 前回调 子类可覆盖做相应初始化工作，一般做样式处理
   * 
   * @param sheet
   *          当前处理的
   */
  protected void beforeSheetFill(Sheet sheet, Map<String, Object> params, StringBuilder returnMsg) {

  }

  /**
   * 填充sheet 完成后回调 子类可覆盖实现批量插入相关收尾工作
   * 
   * @param sheet
   *          当前处理的
   */
  protected void endSheetFill(Sheet sheet, Map<String, Object> params, StringBuilder returnMsg) {

  }

  /**
   * 导出数据放置的sheet 名称
   * 
   * @return sheet 名称，如果使用模板导出，为空则遍历所有sheet
   */
  protected String sheetname() {
    return null;
  }

  /**
   * 填充excel 数据 实现自动判定当前sheet 是否是自己需要填充的页
   * 
   * @param sheet
   *          表页
   * @param params
   *          参数
   * @param resultMsg
   *          返回消息
   */
  protected void fillSheetData(Sheet sheet, Map<String, Object> params, StringBuilder resultMsg) {

  }

  /**
   * 批量导入
   * 
   * @param is
   *          要导入的execl文件流， 由来源负责关闭文件流
   * @param params
   *          ，导入需要的附加参数
   */

  @Override
  // @Transactional
  public void importData(InputStream is, ExcelType excelType, Map<String,
      Object> params,
      final StringBuilder resultMsg) {
    Workbook workbook = null;
    try {
      workbook = ExcelUtils.createWorkBook(is, excelType);
      int sheetIndex = sheetIndex();
      if (params == null) {
        params = new HashMap<String, Object>();
      }
      params.put(TITLE_PARAMS, new HashSet<ExcelTitle>());
      Sheet sheet;
      if (sheetIndex > 0) {
        sheet = workbook.getSheetAt(0);
        logger.debug("parse sheet {} start!", sheetIndex);
        beforeSheetParse(sheet, params, resultMsg);
        parseSheet(sheet, params, resultMsg);
        endSheetParse(sheet, params, resultMsg);
        logger.debug("parse sheet {} end!", sheetIndex);
      } else {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) { // 获取每个Sheet表
          sheet = workbook.getSheetAt(i);
          boolean sheetHidden = workbook.isSheetHidden(i);
          if (!sheetHidden) {
            logger.debug("parse sheet {} start!", i);
            beforeSheetParse(sheet, params, resultMsg);
            parseSheet(sheet, params, resultMsg);
            endSheetParse(sheet, params, resultMsg);
            logger.debug("parse sheet {} end!", i);
          }
        }
      }
    } catch (IOException e) {
      resultMsg.append("excel io异常：").append(e.getMessage());
      logger.error("read excel file failed", e);
    }
  }

  protected void parseSheet(Sheet sheet, Map<String, Object> params, StringBuilder resultMsg) {
    if (sheet.getLastRowNum() < titleLine()) {
      write(resultMsg, "没有待批量注册的数据！");
      logger.info("当前页没有待批量注册的数据");
      return;
    }

    Map<ExcelTitle, Column> headers = parseExcelHeader(sheet, params);
    if (!checkExcelStruct(headers, params, resultMsg)) {
      logger.info("当前页表头结构不符合导入要求");
      return;
    }

    Row row = null;
    for (int i = titleLine() + 1; i <= sheet.getLastRowNum(); i++) {
      row = sheet.getRow(i);
      if (row != null) {
        Map<ExcelTitle, String> rowValueMap = new HashMap<ExcelTitle, String>();
        if (!readRows(row, headers, rowValueMap, params, resultMsg)) {
          continue;
        }

        parseRow(rowValueMap, params, row, resultMsg);
      }
    }
  }

  /**
   * 解析excel 头部
   * 
   * @param sheet
   *          表页
   * @param params
   *          上下文参数
   * @return 表头map
   */
  protected Map<ExcelTitle, Column> parseExcelHeader(Sheet sheet, Map<String, Object> params) {
    if (params.get(TITLE_PARAMS) == null) {
      params.put(TITLE_PARAMS, new HashSet<ExcelTitle>());
    }
    Row row = sheet.getRow(titleLine());
    int columnNum = row.getLastCellNum();
    Map<ExcelTitle, Column> headerIndexMap = new HashMap<ExcelTitle, Column>();
    Set<ExcelTitle> requiredMap = new HashSet<ExcelTitle>();
    for (int index = 0; index < columnNum; index++) {
      String name = ExcelUtils.getCellValue(row.getCell(index));
      if (StringUtils.isBlank(name)) {
        continue;
      }
      TitleWrapper eh = map(name, params);
      if (eh != null) {
        Column cn = Column.newInstance(eh.isRequired(), name, index);
        if (eh.getRealTitle() != null) {
          headerIndexMap.put(eh.getRealTitle(), cn);
          if (eh.isRequired()) {
            requiredMap.add(eh.getRealTitle());
          }
        } else {
          headerIndexMap.put(eh, cn);
        }

      }
    }

    params.put(REQUIRED_HEADERS, requiredMap);

    return headerIndexMap;
  }

  /**
   * 读取行数据，存储到 returnValue 中
   * 
   * @param row
   *          要读取的行
   * @param headerIndexMap
   *          excel 列表头结构
   * @param returnValue
   *          用于存储行数据
   * @param returnMsg
   *          用于存储错误信息
   * @return true if row read success
   */
  private boolean readRows(Row row, Map<ExcelTitle, Column> headerIndexMap, Map<ExcelTitle, String> returnValue,
      Map<String, Object> params, StringBuilder returnMsg) {
    @SuppressWarnings("unchecked")
    Set<ExcelTitle> requiredMap = (Set<ExcelTitle>) params.get(REQUIRED_HEADERS);
    for (ExcelTitle eh : getTitles(params)) {
      Column cn = headerIndexMap.get(eh);
      if (cn != null) {
        String v = ExcelUtils.getCellValue(row.getCell(cn.getIndex()));
        if (StringUtils.isEmpty(v) && (cn.getRequired() || requiredMap.contains(eh))) {
          returnMsg.append("数据错误：第").append(row.getRowNum() + 1).append("行中的").append(cn.getName())
              .append("不能为空,已自动忽略忽略该行;<br>");
          logger.warn("数据错误：第{}行中的{}不能为空,已自动忽略忽略该行。", row.getRowNum() + 1,
              cn.getName());
          return false;
        } else if (v.length() > eh.length()) { // 校验长度
          returnMsg.append("数据警告：第").append(row.getRowNum() + 1).append("行中的").append(cn.getName()).append("内容不得超过")
              .append(eh.length()).append("个字符,已自动截取;<br>");
          v = v.substring(0, eh.length());
          logger.info("数据错误：第{}行中的{}内容不得超过{}个字符,已自动截取。", row.getRowNum() + 1,
              cn.getName(), eh.length());
        }

        returnValue.put(eh, v);
      }
    }

    for (ExcelTitle title : headerIndexMap.keySet()) {
      if (title instanceof TitleWrapper) {
        TitleWrapper wrapper = (TitleWrapper) title;
        if (wrapper.getRealTitle() == null) { // 处理自定义字段
          Column cn = headerIndexMap.get(title);
          String v = ExcelUtils.getCellValue(row.getCell(cn.getIndex()));
          if (StringUtils.isEmpty(v) && cn.getRequired()) {
            returnMsg.append("数据错误：第").append(row.getRowNum() + 1).append("行中的").append(cn.getName())
                .append("不能为空,已自动忽略忽略该行;<br>");
            logger.warn("数据错误：第{}行中的{}不能为空,已自动忽略忽略该行。", row.getRowNum() + 1,
                cn.getName());
            return false;
          }
          returnValue.put(title, v);
        }
      }
    }

    return true;
  }

  private TitleWrapper map(final String name, Map<String, Object> params) {
    String sname = prepareTitle(name);
    ExcelTitle header = null;
    for (ExcelTitle eh : getTitles(params)) {
      if (eh.getMapNames().contains(sname)) {
        header = eh;
        break;
      }
    }
    return new TitleWrapper(sname, isRequiredColumn(name), header);
  }

  /**
   * 检查excel 结构是否正确，根据必须的列头判定
   * 
   * @param headerIndexMap
   *          head title map
   * @return 结构正确 true, 错误为false
   */
  private boolean checkExcelStruct(Map<ExcelTitle, Column> headerIndexMap, Map<String, Object> params,
      StringBuilder returnMsg) {
    boolean rs = true;
    for (ExcelTitle eh : requiedHeaders(params)) {
      if (headerIndexMap.get(eh) == null) {
        returnMsg.append("Excel 缺少必须列： ").append(eh.getMapNames().get(0)).append("<br/>");
        logger.warn("Excel 缺少必须列：{}", eh.getMapNames().get(0));
        rs = false;
      }
    }
    return rs;
  }

  protected List<ExcelTitle> requiedHeaders(Map<String, Object> params) {
    List<ExcelTitle> requiredTitles = new ArrayList<ExcelTitle>();
    Set<String> rnms = requireColmnNames();
    for (ExcelTitle eh : getTitles(params)) {
      if (eh.isRequired() || configRequiredColumn(eh.getMapNames(), rnms)) {
        requiredTitles.add(eh);
      }
    }

    return requiredTitles;
  }

  private boolean configRequiredColumn(List<String> names, Set<String> rnms) {
    if (names != null) {
      for (String name : names) {
        if (rnms.contains(name)) {
          return true;
        }
      }
    }

    return false;
  }

  /**
   * 读取自定义配置的必须列名称。 默认配置名称为 ： {simpleClassName}.rqColumns 多个配置需以 "," 分隔
   * 
   * @return set of required column name.
   */
  protected Set<String> requireColmnNames() {
    Set<String> rnms = new HashSet<String>();
    String key = requiredColumnKeyName();
    // String rqmstr = ConfigUtils.readConfig(key, "");
    // if (StringUtils.isNotBlank(rqmstr)) {
    // // logger.debug("use custom required column config。 [{} : {}]", key,
    // // rqmstr);
    // String[] rnmarr = StringUtils.split(rqmstr, ",");
    // for (String nm : rnmarr) {
    // if (StringUtils.isNotBlank(nm)) {
    // rnms.add(nm.trim());
    // }
    // }
    // }

    return rnms;
  }

  /**
   * 配置名称key
   * 
   * @return column key name
   */
  protected String requiredColumnKeyName() {
    return getClass().getSimpleName() + "." + "batch.rqColumns";
  }

  private void write(StringBuilder resultMsg, String msg) {
    if (resultMsg != null) {
      resultMsg.append(msg);
    }
  }

  @SuppressWarnings("unchecked")
  protected Set<ExcelTitle> getTitles(Map<String, Object> params) {
    Set<ExcelTitle> titles = (Set<ExcelTitle>) params.get(TITLE_PARAMS);
    if (titles.size() == 0 && titles() != null) {
      titles = new LinkedHashSet<ExcelTitle>();
      titles.addAll(Arrays.asList(titles()));
    }

    return titles;
  }

  @SuppressWarnings("unchecked")
  protected void addTitle(ExcelTitle title, Map<String, Object> params) {
    Set<ExcelTitle> titles = (Set<ExcelTitle>) params.get(TITLE_PARAMS);
    titles.add(title);
  }

  /**
   * 获取头部信息
   * 
   * @return ExcelTitle[]
   */
  protected abstract ExcelTitle[] titles();

  /**
   * 解析sheet 前回调 子类可覆盖做相应初始化工作
   * 
   * @param sheet
   *          当前处理的
   */
  protected void beforeSheetParse(Sheet sheet, Map<String, Object> params, StringBuilder returnMsg) {

  }

  /**
   * 解析sheet 完成后回调 子类可覆盖实现批量插入相关收尾工作
   * 
   * @param sheet
   *          当前处理的
   */
  protected void endSheetParse(Sheet sheet, Map<String, Object> params, StringBuilder returnMsg) {

  }

  /**
   * 获取当前要读取的sheet
   * 
   * @return 当前sheet,小于0 则遍历读取excel 所有sheet
   */
  public abstract int sheetIndex();

  /**
   * 获取excel模板标题所在行 用于解析标题行，确定excel 模板正确性
   * 
   * @return title line index
   */
  public abstract int titleLine();

  /**
   * 处理解析好的单行数据
   * 
   * @param rowValueMap
   *          row data
   * @param params
   *          context
   * @param row
   *          sheet row
   * @param returnMsg
   *          return message
   */
  protected abstract void parseRow(Map<ExcelTitle, String> rowValueMap, Map<String, Object> params, Row row,
      StringBuilder returnMsg);

  /**
   * 处理标题内容，去掉空白字符，去掉* 号 子类覆盖做特殊处理
   * 
   * @param name
   *          column title
   * @return tiltle without '*'
   */
  protected String prepareTitle(String name) {
    return name != null ? name.replaceAll("\\*", "").replaceAll("\\s", "").replaceAll("\\（必填\\）", "") : "";
  }

  /**
   * 判定模板是否设置标题列为必填
   * 
   * @param columnName
   *          columnName
   * @return true if is require
   */
  protected boolean isRequiredColumn(String columnName) {
    if (columnName != null) {
      return columnName.contains("*");
    }
    return false;
  }

  public static class TitleWrapper implements ExcelTitle {

    private final boolean requierd;

    private final ExcelTitle realTitle;

    private final String name;

    TitleWrapper(final String name, final boolean required, ExcelTitle realTitle) {
      this.requierd = required;
      this.realTitle = realTitle;
      this.name = name;
    }

    @Override
    public List<String> getMapNames() {
      return realTitle == null ? Arrays.asList(name) : realTitle.getMapNames();
    }

    @Override
    public boolean isRequired() {
      return (realTitle != null && realTitle.isRequired()) || requierd;
    }

    @Override
    public int length() {
      return realTitle == null ? Integer.MAX_VALUE : realTitle.length();
    }

    /**
     * Getter method for property <tt>realTitle</tt>.
     * 
     * @return property value of realTitle
     */
    public ExcelTitle getRealTitle() {
      return realTitle;
    }

    @Override
    public boolean equals(Object obj) {
      if (realTitle != null) {
        return realTitle.equals(obj);
      }
      return super.equals(obj);
    }

    @Override
    public int hashCode() {
      if (realTitle != null) {
        return realTitle.hashCode();
      }
      return super.hashCode();
    }
  }

}
