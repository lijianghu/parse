/**
 * Mainbo.com Inc.
 * Copyright (c) 2015-2017 All Rights Reserved.
 */

package cn.timesaving.core.export;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import cn.timesaving.core.export.ExcelUtils.ExcelType;

/**
 * <pre>
 * 批量处理基础服务类
 * </pre>
 *
 * @author tmser
 * @version $Id: BatchImportService.java, v 1.0 2016年8月30日 下午4:42:06 tmser Exp $
 */
public interface BatchService {

  /**
   * 批量导入
   * 
   * @param is
   *          要导入的文件流， 由来源负责关闭文件流
   * @param params
   *          导入附加参数
   * @param resultMsg
   *          ，导入消息记录器
   */
  void importData(InputStream is, ExcelType excelType, Map<String, Object> params, StringBuilder resultMsg);

  /**
   * 批量导出服务
   * 
   * @param os
   *          输出流,
   * @param template
   *          要使用的导出模板，允许为null
   * @param resultMsg
   *          ，导出消息记录器
   */
  void exportData(OutputStream os, Map<String, Object> params, StringBuilder resultMsg, File template);

  interface ExcelTitle {
    /**
     * 标题名称列表，即允许一列使用不同的标题
     * 
     * @return title names
     */
    List<String> getMapNames();

    /**
     * 当前列是否必须
     * 
     * @return true if required
     */
    boolean isRequired();

    /**
     * 内容长度限制
     * 
     * @return lenth
     */
    int length();
  }

  /**
   * 
   * <pre>
   * 数据列
   * </pre>
   *
   * @author tmser
   * @version $Id: ExcelImportService.java, v 1.0 2016年8月30日 下午5:32:06 tmser Exp
   *          $
   */
  static class Column {
    static Column newInstance(Boolean required, String name, Integer index) {
      return new Column(required, name, index);
    }

    public Column(Boolean required, String name, Integer index) {
      this.name = name;
      this.index = index;
      this.required = required;
    }

    private Integer index;

    private String name;

    private Boolean required = Boolean.FALSE;

    public Integer getIndex() {
      return index;
    }

    public String getName() {
      return name;
    }

    /**
     * Getter method for property <tt>required</tt>.
     * 
     * @return property value of required
     */
    public Boolean getRequired() {
      return required;
    }

  }
}
