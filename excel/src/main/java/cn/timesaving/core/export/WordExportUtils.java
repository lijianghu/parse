package cn.timesaving.core.export;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

/**
 * 适用于word 2007
 * poi 版本 3.7
 */
public class WordExportUtils {

  /**
   * 根据指定的参数值、模板，生成 word 文档
   * 
   * @param param
   *          需要替换的变量
   * @param template
   *          模板
   */
  public static CustomXWPFDocument generateWord(Map<String, Object> param, String template) {
    CustomXWPFDocument doc = null;
    try {
      OPCPackage pack = POIXMLDocument.openPackage(template);
      doc = new CustomXWPFDocument(pack);
      if (param != null && param.size() > 0) {

        // 处理段落
        List<XWPFParagraph> paragraphList = doc.getParagraphs();
        processParagraphs(paragraphList, param, doc);

        // 处理表格
        Iterator<XWPFTable> it = doc.getTablesIterator();
        while (it.hasNext()) {
          XWPFTable table = it.next();
          List<XWPFTableRow> rows = table.getRows();
          for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
              List<XWPFParagraph> paragraphListTable = cell.getParagraphs();
              processParagraphs(paragraphListTable, param, doc);
            }
          }
        }
      }
    } catch (Exception e) {
      e.printStackTrace();
    }
    return doc;
  }

  /**
   * 处理段落
   * 
   * @param paragraphList
   */
  public static void processParagraphs(List<XWPFParagraph> paragraphList, Map<String, Object> param,
      CustomXWPFDocument doc) {
    if (paragraphList != null && paragraphList.size() > 0) {
      for (XWPFParagraph paragraph : paragraphList) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (XWPFRun run : runs) {
          String text = run.getText(0);
          if (text != null) {
            boolean isSetText = false;
            for (Entry<String, Object> entry : param.entrySet()) {
              String key = entry.getKey();
              if (text.indexOf(key) != -1) {
                isSetText = true;
                Object value = entry.getValue();
                if (value instanceof String) {// 文本替换
                  text = text.replace(key, value.toString());
                } else if (value instanceof Map) {// 图片替换
                  text = text.replace(key, "");
                  Map pic = (Map) value;
                  int width = Integer.parseInt(pic.get("width").toString());
                  int height = Integer.parseInt(pic.get("height").toString());
                  int picType = getPictureType(pic.get("type").toString());
                  byte[] byteArray = (byte[]) pic.get("content");
                  ByteArrayInputStream byteInputStream = new ByteArrayInputStream(byteArray);
                  try {
                    int ind = doc.addPicture(byteInputStream, picType);
                    doc.createPicture(ind, width, height, paragraph);
                  } catch (Exception e) {
                    e.printStackTrace();
                  }
                }
              }
            }
            if (isSetText) {
              run.setText(text, 0);
            }
          }
        }
      }
    }
  }

  /**
   * 根据图片类型，取得对应的图片类型代码
   * 
   * @param picType
   * @return int
   */
  private static int getPictureType(String picType) {
    int res = CustomXWPFDocument.PICTURE_TYPE_PICT;
    if (picType != null) {
      if (picType.equalsIgnoreCase("png")) {
        res = CustomXWPFDocument.PICTURE_TYPE_PNG;
      } else if (picType.equalsIgnoreCase("dib")) {
        res = CustomXWPFDocument.PICTURE_TYPE_DIB;
      } else if (picType.equalsIgnoreCase("emf")) {
        res = CustomXWPFDocument.PICTURE_TYPE_EMF;
      } else if (picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")) {
        res = CustomXWPFDocument.PICTURE_TYPE_JPEG;
      } else if (picType.equalsIgnoreCase("wmf")) {
        res = CustomXWPFDocument.PICTURE_TYPE_WMF;
      }
    }
    return res;
  }

  /**
   * 将输入流中的数据写入字节数组
   * 
   * @param in
   * @return
   */
  public static byte[] inputStream2ByteArray(InputStream in, boolean isClose) {
    byte[] byteArray = null;
    try {
      int total = in.available();
      byteArray = new byte[total];
      in.read(byteArray);
    } catch (IOException e) {
      e.printStackTrace();
    } finally {
      if (isClose) {
        try {
          in.close();
        } catch (Exception e2) {
          System.out.println("关闭流失败");
        }
      }
    }
    return byteArray;
  }

  public static class CustomXWPFDocument extends XWPFDocument {
    public CustomXWPFDocument(InputStream in) throws IOException {
      super(in);
    }

    public CustomXWPFDocument() {
      super();
    }

    public CustomXWPFDocument(OPCPackage pkg) throws IOException {
      super(pkg);
    }

    /**
     * @param id
     * @param width
     *          宽
     * @param height
     *          高
     * @param paragraph
     *          段落
     */
    public void createPicture(int id, int width, int height, XWPFParagraph paragraph) {
      final int EMU = 9525;
      width *= EMU;
      height *= EMU;
      String blipId = getAllPictures().get(id).getPackageRelationship().getId();
      CTInline inline = paragraph.createRun().getCTR().addNewDrawing().addNewInline();
      String picXml = ""
          + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
          + "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
          + "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
          + "         <pic:nvPicPr>" + "            <pic:cNvPr id=\""
          + id
          + "\" name=\"Generated\"/>"
          + "            <pic:cNvPicPr/>"
          + "         </pic:nvPicPr>"
          + "         <pic:blipFill>"
          + "            <a:blip r:embed=\""
          + blipId
          + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"
          + "            <a:stretch>"
          + "               <a:fillRect/>"
          + "            </a:stretch>"
          + "         </pic:blipFill>"
          + "         <pic:spPr>"
          + "            <a:xfrm>"
          + "               <a:off x=\"0\" y=\"0\"/>"
          + "               <a:ext cx=\""
          + width
          + "\" cy=\""
          + height
          + "\"/>"
          + "            </a:xfrm>"
          + "            <a:prstGeom prst=\"rect\">"
          + "               <a:avLst/>"
          + "            </a:prstGeom>"
          + "         </pic:spPr>"
          + "      </pic:pic>"
          + "   </a:graphicData>" + "</a:graphic>";

      inline.addNewGraphic().addNewGraphicData();
      XmlToken xmlToken = null;
      try {
        xmlToken = XmlToken.Factory.parse(picXml);
      } catch (XmlException xe) {
        xe.printStackTrace();
      }
      inline.set(xmlToken);

      inline.setDistT(0);
      inline.setDistB(0);
      inline.setDistL(0);
      inline.setDistR(0);

      CTPositiveSize2D extent = inline.addNewExtent();
      extent.setCx(width);
      extent.setCy(height);

      CTNonVisualDrawingProps docPr = inline.addNewDocPr();
      docPr.setId(id);
      docPr.setName("头像" + id);
      docPr.setDescr("头像");
    }
  }
}
