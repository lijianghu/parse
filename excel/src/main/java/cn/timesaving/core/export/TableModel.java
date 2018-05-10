package cn.timesaving.core.export;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class TableModel {
  private static Logger logger = LoggerFactory.getLogger(TableModel.class);
  private String name;
  private List<String> headers = new ArrayList<String>();
  private List<String> headerNames = new ArrayList<String>();
  private List data;

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public void addHeaders(String... header) {
    if (header == null) {
      return;
    }

    for (String text : header) {
      if (text == null) {
        continue;
      }

      headers.add(text);
    }
  }

  public void addHeaderNames(String... headerNames) {
    if (headerNames == null) {
      return;
    }

    for (String text : headerNames) {
      if (text == null) {
        continue;
      }

      this.headerNames.add(text);
    }
  }

  public void setData(List data) {
    this.data = data;
  }

  public int getHeaderCount() {
    return headers.size();
  }

  public int getDataCount() {
    return data.size();
  }

  public String getHeader(int index) {
    return headers.get(index);
  }

  public String getHeaderName(int index) {
    return CollectionUtils.isNotEmpty(headerNames) ? headerNames.get(index) : headers.get(index);
  }

  public String getValue(int i, int j) {
    String header = getHeader(j);
    Object object = data.get(i);

    if (object instanceof Map) {
      return this.getValueFromMap(object, header);
    } else {
      return this.getValueReflect(object, header);
    }
  }

  public String getValueReflect(Object instance, String fieldName) {
    try {
      String methodName = ReflectUtils.getGetterMethodName(instance, fieldName);
      Object value = ReflectUtils.getMethodValue(instance, methodName);

      return formatData(value);
    } catch (Exception ex) {
      logger.info("error", ex);

      return "";
    }
  }

  private String formatData(Object data) {
    if (data == null) {
      return "";
    }
    String value = "";
    if (data instanceof Date) {
      value = DateFormatUtils.format((Date) data, "yyyy-MM-dd HH:mm");
    } else {
      value = data.toString();
    }

    return value;

  }

  /**
   * <p>
   * </p>
   * Getter method for property <tt>headerNames</tt>.
   *
   * @return headerNames List<String>
   */
  public List<String> getHeaderNames() {
    return headerNames;
  }

  /**
   * <p>
   * </p>
   * Setter method for property <tt>headerNames</tt>.
   *
   * @param headerNames
   *          List<String> value to be assigned to property headerNames
   */
  public void setHeaderNames(List<String> headerNames) {
    this.headerNames = headerNames;
  }

  public String getValueFromMap(Object instance, String fieldName) {
    Map<String, Object> map = (Map<String, Object>) instance;
    Object value = map.get(fieldName);

    return formatData(value);
  }
}
