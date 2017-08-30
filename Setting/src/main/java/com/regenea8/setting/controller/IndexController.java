package com.regenea8.setting.controller;

import java.net.InetAddress;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;
import javax.inject.Inject;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import com.regenea8.setting.service.IndexService;
import com.regenea8.setting.util.InternetProtocolUtil;
import com.regenea8.setting.vo.TestVO;


@Controller
@RequestMapping("/*")
public class IndexController {

	private static final Logger logger = LoggerFactory.getLogger(IndexController.class);

	@Inject
	private IndexService service;

	@Resource(name = "imgUploadPath")
	private String uploadPath;

	@RequestMapping(value = "/index", method = { RequestMethod.GET, RequestMethod.POST})
	public void index(HttpServletRequest request, Model model) throws Exception {

		logger.info("-------------start index [Connect IP : " + InternetProtocolUtil.getClientIp(request) + "]");

		List list = null;

		try {

			list = service.testList();
			
		} catch (Exception e) {
			e.printStackTrace();
		}

		model.addAttribute("list", list);
		

		logger.info("---------------end index [Connect IP : " + InternetProtocolUtil.getClientIp(request) + "]");
	}
	
	@ResponseBody
	@RequestMapping(value = "/test", method = { RequestMethod.GET, RequestMethod.POST})
	public String download(HttpServletRequest request, TestVO vo) throws Exception {

		logger.info("-------------start download [Connect IP : " + InternetProtocolUtil.getClientIp(request) + "]");

		String result = "";
		
		try {
		
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		logger.info("---------------end download [Connect IP : " + InternetProtocolUtil.getClientIp(request) + "]");
		
		return result;
	}
}
