package com.dexter.parameter.controller;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.dexter.conn.Conn;
import com.dexter.parameter.entity.Parameter;

public class ParameterController {
	/*
	 * 连接数据库获取参数列表
	 * category: 查询的参数
	 */
	public List<Parameter> getParametersFromDBByCategory(String category){
		List<Parameter> params = new ArrayList<Parameter>();
		try{
			Connection conn = Conn.getConnectionForJDBC();
			Statement statement = conn.createStatement();
			String sql = "select name, value from cs_sys_parameter where category = '" + category + "'";
					
			ResultSet rs = statement.executeQuery(sql);
			Parameter p = null;
			while(rs.next()){
				p = new Parameter();
				p.setName(rs.getString(1));
				p.setValue(rs.getString(2));
				System.out.println(p.getName() + ":" + p.getValue());
				params.add(p);
			}
			
		}catch(Exception e){
			e.printStackTrace();
		}
		return params;
	}
	
	/*
	 * 连接数据库获取参数列表
	 * category: 查询的参数
	 */
	public Map<String, String> getMapFromDBByCategory(String category){
		Map<String, String> map = new HashMap<String, String>();
		try{
			Connection conn = Conn.getConnectionForJDBC();
			Statement statement = conn.createStatement();
			String sql = "select name, value from cs_sys_parameter where category = '" + category + "'";
					
			ResultSet rs = statement.executeQuery(sql);
			while(rs.next()){
				//TODO 这样若是有相同名称的数据就会被覆盖。
				map.put(rs.getString(1), rs.getString(2));
			}
		}catch(Exception e){
			e.printStackTrace();
		}
		return map;
	}
	
	public static void main(String[] args) {
		ParameterController pc = new ParameterController();
		//pc.getParametersFromDBByCategory("education_degree");
		pc.getMapFromDBByCategory("education_degree");
	}
	
}
