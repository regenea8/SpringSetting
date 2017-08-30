package com.regenea8.setting.daoImpl;

import java.util.List;

import javax.inject.Inject;

import org.apache.ibatis.session.SqlSession;
import org.springframework.stereotype.Repository;

import com.regenea8.setting.dao.IndexDAO;

@Repository
public class IndexDAOImpl implements IndexDAO {
	
	@Inject
	private SqlSession session;

	private static String namespace = "com.regenea8.setting.IndexMapper";
	
	@Override
	public List testList() throws Exception {
		return session.selectList(namespace + ".testList");
	}
}
