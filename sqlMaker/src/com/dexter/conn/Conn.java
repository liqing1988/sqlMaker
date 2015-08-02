package com.dexter.conn;

import java.io.ObjectInputStream.GetField;
import java.sql.Connection;
import java.sql.DriverManager;
import java.util.ResourceBundle;

public class Conn {
	public static Connection getConnectionForJDBC() {
		java.sql.Connection conn =null;
		try {
			String fileName = "application";
			ResourceBundle resource = ResourceBundle.getBundle(fileName);
			String driver = resource.getString("jdbc.driver");
			String url = resource.getString("jdbc.url");
			String username = resource.getString("jdbc.username");
			String password = resource.getString("jdbc.password");
			//System.out.println(driver+":"+url+":"+username+":"+password);
			Class.forName(driver);
			conn = DriverManager.getConnection(url, username, password);
			conn.setAutoCommit(false);
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("连接数据库失败...");
		}
		return conn;
	}
	
	/*
	 * just for test
	 */
	public static void main(String[] args) {
		Conn conn = new Conn();
		conn.getConnectionForJDBC();
	}
}
