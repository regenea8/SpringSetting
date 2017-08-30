package com.regenea8.setting.serviceImpl;

import java.util.List;

import javax.inject.Inject;

import org.springframework.stereotype.Service;

import com.regenea8.setting.dao.IndexDAO;
import com.regenea8.setting.service.IndexService;

@Service
public class IndexServiceImpl implements IndexService {
	
	@Inject
	private IndexDAO dao;
	
	@Override
	public List testList() throws Exception {
		return dao.testList();
	}

}
