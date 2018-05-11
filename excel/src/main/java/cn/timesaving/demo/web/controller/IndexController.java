package cn.timesaving.demo.web.controller;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;

/**
 * <pre>
 * 测试
 * </pre>
 *
 * @author ljh
 * @version $Id: IndexController.java, v 1.0 2018年5月11日 下午4:12:16 ljh Exp $
 */
@Controller
@RequestMapping
public class IndexController {

  private static Logger logger = LoggerFactory.getLogger(TestDormitoryNumImportSetvice.class);

  @Autowired
  private TestDormitoryNumImportSetvice testDormitoryNumImportSetvice;

  @RequestMapping("/index")
  public String list(HttpServletRequest request) {

    return "index";
  }

  @RequestMapping("/import")
  public void testimport(MultipartFile file) {
    try {
      testDormitoryNumImportSetvice.import_(file);
    } catch (Exception e) {
      logger.info("", e);
    }
  }

  @RequestMapping("/export")
  public void testexport(HttpServletRequest request, HttpServletResponse response) {
    testDormitoryNumImportSetvice.export(request, response);

  }
}