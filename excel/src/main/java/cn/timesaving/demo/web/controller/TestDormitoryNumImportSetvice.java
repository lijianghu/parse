package cn.timesaving.demo.web.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.context.ContextLoader;
import org.springframework.web.multipart.MultipartFile;

import cn.timesaving.core.export.ExcelBatchService;
import cn.timesaving.core.export.ExcelUtils;
import cn.timesaving.core.export.ExcelUtils.ExcelType;

@Service
public class TestDormitoryNumImportSetvice extends ExcelBatchService {

  private static Logger logger = LoggerFactory.getLogger(TestDormitoryNumImportSetvice.class);

  private final Integer headline = 2;
  private final String sheetName = "Sheet1";

  /**
   * 导出
   * 
   * @param request
   * @param response
   */
  public void export(HttpServletRequest request, HttpServletResponse response) {
    Workbook workbook = null;

    String fileName = "批量导入宿舍号";
    String filepath = "";
    filepath = "/importtemplate/num_template.xlsx";
    String realPath = ContextLoader.getCurrentWebApplicationContext().getServletContext().getRealPath(filepath);
    workbook = ExcelUtils.createWorkBook(realPath);
    if (workbook == null) {
      returnErrMsg(filepath, response);
      return;
    }
    Sheet sheet = workbook.getSheet(sheetName);
    Map<String, Object> params = new HashMap<String, Object>();
    params.put(TITLE_PARAMS, new HashSet<ExcelTitle>());
    Map<ExcelTitle, Column> headers = parseExcelHeader(sheet, params);
    initSheetTeplate(workbook, sheet, headers);
    try {
      String ext = filepath.substring(filepath.lastIndexOf("."));
      response.reset();
      response.setContentType("application/x-msdownload");
      response.setHeader("Content-Disposition", "attachment; filename="
          + new String(fileName.getBytes("gb2312"), "ISO-8859-1") + ext);
      workbook.write(response.getOutputStream());
    } catch (Exception e) {
      logger.error("", e);
    }
  }

  /**
   * 初始化模板数据
   * 
   * @param workbook
   * @param sheet
   * @param headers
   */
  private void initSheetTeplate(Workbook workbook, Sheet sheet, Map<ExcelTitle, Column> headers) {
    Column cn = null;
    if ((cn = headers.get(ExcelHeader.FLOOR)) != null) { // 初始化宿舍楼名称
      List<DormitoryFloor> floorList = getDormitoryFloorList();
      String[] floorNameArray = new String[floorList.size()];
      for (int i = 0; i < floorList.size(); i++) {
        floorNameArray[i] = floorList.get(i).getFloorName();
      }
      Integer floorColumnIndex = cn.getIndex();
      DataValidation setDataValidationName3 = ExcelUtils.createDataValidation(
          sheet, floorNameArray, headline + 1,
          300, floorColumnIndex, floorColumnIndex);
      sheet.addValidationData(setDataValidationName3);
    }
  }

  /**
   * for test
   * 
   * @return
   */
  private List<DormitoryFloor> getDormitoryFloorList() {
    List<DormitoryFloor> testlist = new ArrayList<DormitoryFloor>();
    for (int i = 0; i < 10; i++) {
      DormitoryFloor df = new DormitoryFloor();
      df.setCreateTime(new Date());
      df.setFloorName("宿舍楼" + i + "号[编码：" + UUID.randomUUID() + "]");
      df.setId(123456L);
      df.setOrgId("orgId001");
      df.setUserId(12345678L);
      testlist.add(df);
    }
    return testlist;
  }

  private void returnErrMsg(String path, HttpServletResponse response) {
    response.setHeader("Content-type", "text/html;charset=UTF-8");
    StringBuilder rb = new StringBuilder(path).append(" 不存在!");
    try {
      response.getWriter().write(rb.toString());
    } catch (IOException e) {
      logger.error("", e);
    }
  }

  /**
   * @return
   * @see com.mossle.core.export.ExcelBatchService#titles()
   */
  @Override
  protected ExcelTitle[] titles() {
    return ExcelHeader.values();
  }

  /**
   * @return
   * @see com.mossle.core.export.ExcelBatchService#sheetIndex()
   */
  @Override
  public int sheetIndex() {
    return 0;
  }

  /**
   * @return
   * @see com.mossle.core.export.ExcelBatchService#titleLine()
   */
  @Override
  public int titleLine() {
    return headline;
  }

  /**
   * 导入
   * 
   * @param file
   * @return
   */
  public StringBuilder import_(MultipartFile file) {
    Map<String, Object> params = new HashMap<String, Object>();
    StringBuilder resultStr = new StringBuilder(); // 反馈信息
    try {
      resultStr.append("<strong>==》开始导入数据，请等待...</strong><br>");
      this.importData(file.getInputStream(),
          ExcelType.parseExcelType(file.getOriginalFilename()), params, resultStr);
      resultStr.append("<br><strong>《==数据导入结束。</strong>");
    } catch (IOException e) {
      resultStr.append("批量导入失败：").append(e.getMessage());
      logger.error("import dormitory num file failed", e);
    }
    return resultStr;
  }

  // 初始化一些数据
  @Override
  protected void beforeSheetParse(Sheet sheet, Map<String, Object> params,
      StringBuilder returnMsg) {
    Map<String, Long> floorMap = new HashMap<String, Long>(); // 机构下所有的宿舍楼
    List<DormitoryFloor> floorList = getDormitoryFloorList();
    for (int i = 0; i < floorList.size() - 2; i++) {// 这里让初始化少两条数据，让下面程序发出警告
      floorMap.put(floorList.get(i).getFloorName(), floorList.get(i).getId());
    }
    params.put("floorMap", floorMap);
  }

  /**
   * @param rowValueMap
   * @param params
   * @param row
   * @param returnMsg
   * @see com.mossle.core.export.ExcelBatchService#parseRow(java.util.Map,
   *      java.util.Map, org.apache.poi.ss.usermodel.Row,
   *      java.lang.StringBuilder)
   */
  @Override
  protected void parseRow(Map<ExcelTitle, String> rowValueMap, Map<String, Object> params, Row row,
      StringBuilder sb) {
    if (row != null) {
      int i = row.getRowNum() + 1;
      Map<String, Long> floorMap = (Map<String, Long>) params.get("floorMap");
      String floorName = rowValueMap.get(ExcelHeader.FLOOR);
      // 数据校验
      if (!checkCulumn(i, rowValueMap, sb, params)) {
        return;
      }
      // 这里进行业务操作，当前解析行数据
      logger.info("-----------业务操作--------floorName：" + floorName);
    }
  }

  /**
   * @param i
   * @param rowValueMap
   * @param returnMsg
   * @param params
   * @return
   */
  private boolean checkCulumn(int i, Map<ExcelTitle, String> rowValueMap, StringBuilder sb,
      Map<String, Object> params) {
    Map<String, Long> floorMap = (Map<String, Long>) params.get("floorMap");
    String floorName = rowValueMap.get(ExcelHeader.FLOOR);
    String numName = rowValueMap.get(ExcelHeader.NUM);
    if (StringUtils.isEmpty(floorName) || StringUtils.isEmpty(numName)) {
      sb.append("第" + (i) + "行必填项有空值，已忽略此行数据。<br>");
      logger.debug("第{}行必填项有空值，忽略此行数据。", i);
      return false;
    }
    if (StringUtils.isNotBlank(floorName)
        && floorMap.get(floorName) == null) {
      sb.append("第" + (i) + "行：宿舍楼不存在，已忽略此行数据。<br>");
      logger.debug("第" + (i) + "行：宿舍楼不存在，已忽略此行数据", i);
      return false;
    }
    return true;
  }

  /**
   * 判定模板是否设置标题列为必填
   * 
   */
  @Override
  protected boolean isRequiredColumn(String columnName) {
    if (columnName != null) {
      return columnName.contains("必填");
    }
    return super.isRequiredColumn(columnName);
  }

  /**
   * 处理标题内容，去掉空白字符，去掉* 号 子类覆盖做特殊处理
   * 
   */
  @Override
  protected String prepareTitle(String name) {
    name = super.prepareTitle(name);
    return name != null && name.length() > 4 ? name.replaceAll("\\（必填\\）", "")
        : name;
  }

  enum ExcelHeader implements ExcelTitle {
    FLOOR(true, 32, "宿舍楼"),
    NUM(true, 32, "宿舍号");

    private String[] mapNames;

    private boolean required = false;

    private Integer size = null;

    private ExcelHeader(String... header) {
      this.mapNames = header;
    }

    private ExcelHeader(boolean required, String... header) {
      this.mapNames = header;
      this.required = required;
    }

    private ExcelHeader(Integer size, String... header) {
      this.mapNames = header;
      this.size = size;
    }

    private ExcelHeader(boolean required, Integer size, String... header) {
      this.mapNames = header;
      this.required = required;
      this.size = size;
    }

    @Override
    public List<String> getMapNames() {
      return Arrays.asList(mapNames);
    }

    @Override
    public boolean isRequired() {
      return required;
    }

    /**
     * @see com.mainbo.platform.common.service.ExcelImportService.ExcelTitle#size()
     */
    @Override
    public int length() {
      return size;
    }
  }

  // test class
  static class DormitoryFloor {
    private Long id;

    private Date createTime;

    private Long userId;

    private String orgId;

    private String floorName;

    /**
     * <p>
     * </p>
     * Getter method for property <tt>id</tt>.
     *
     * @return id Long
     */
    public Long getId() {
      return id;
    }

    /**
     * <p>
     * </p>
     * Setter method for property <tt>id</tt>.
     *
     * @param id
     *          Long value to be assigned to property id
     */
    public void setId(Long id) {
      this.id = id;
    }

    /**
     * <p>
     * </p>
     * Getter method for property <tt>createTime</tt>.
     *
     * @return createTime Date
     */
    public Date getCreateTime() {
      return createTime;
    }

    /**
     * <p>
     * </p>
     * Setter method for property <tt>createTime</tt>.
     *
     * @param createTime
     *          Date value to be assigned to property createTime
     */
    public void setCreateTime(Date createTime) {
      this.createTime = createTime;
    }

    /**
     * <p>
     * </p>
     * Getter method for property <tt>userId</tt>.
     *
     * @return userId Long
     */
    public Long getUserId() {
      return userId;
    }

    /**
     * <p>
     * </p>
     * Setter method for property <tt>userId</tt>.
     *
     * @param userId
     *          Long value to be assigned to property userId
     */
    public void setUserId(Long userId) {
      this.userId = userId;
    }

    /**
     * <p>
     * </p>
     * Getter method for property <tt>orgId</tt>.
     *
     * @return orgId String
     */
    public String getOrgId() {
      return orgId;
    }

    /**
     * <p>
     * </p>
     * Setter method for property <tt>orgId</tt>.
     *
     * @param orgId
     *          String value to be assigned to property orgId
     */
    public void setOrgId(String orgId) {
      this.orgId = orgId;
    }

    /**
     * <p>
     * </p>
     * Getter method for property <tt>floorName</tt>.
     *
     * @return floorName String
     */
    public String getFloorName() {
      return floorName;
    }

    /**
     * <p>
     * </p>
     * Setter method for property <tt>floorName</tt>.
     *
     * @param floorName
     *          String value to be assigned to property floorName
     */
    public void setFloorName(String floorName) {
      this.floorName = floorName;
    }
  }

}
