package com.dexter.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import net.infotop.util.OperationNoUtil;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.dexter.parameter.controller.ParameterController;

public class ConvertExcel {

	public void convertExcel(String filePath){
		//获取导出文件名
		//转换参数并导出excel
	}
	
	/*
	 * 获取导出文件的路径
	 * filePath：源文件路径
	 * outFilePaths ： 0 文件路径；1 生成的文件名称； 
	 */
	public String[] getOutFilePath(String inFilePath){
		String[] tempPaths = inFilePath.split("/");
		String[] outFilePaths = new String[2];
		String tempPath = "";	//转换后的excel文件路径，含路径
		String tempFileName = "";	//转换后的excel名称
		for(int i = 0; i < (tempPaths.length - 1); i++){
			tempPath += tempPaths[i] + "/";
		}
		//转换后的excel名称为 年月日时分秒_原名称
		SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmmss");
		tempFileName += sdf.format(new Date()) + "_" + tempPaths[tempPaths.length-1];
		
		outFilePaths[0] = tempPath;
		outFilePaths[1] = tempFileName;
		return outFilePaths;
	}
	
	/*
	 * 获取导出SQL文件的路径
	 * filePath：源文件路径
	 * sheetName : 工作薄名称
	 */
	public String getOutFilePath(String inFilePath, String sheetName){
		String[] paths = inFilePath.split("/");
		String tempPath = "";	//转换后的excel文件路径，含路径
		String tempFileName = "";	//转换后的excel名称
		for(int i = 0; i < (paths.length - 1); i++){
			tempPath += paths[i] + "/";
		}
		//转换后的excel名称为 年月日时分秒_原名称
		SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmmss");
		tempFileName += sdf.format(new Date()) + "_" + paths[paths.length-1];
		
		tempPath += tempFileName;
		return tempPath;
	}
	
	/*
	 * 转换内容，导出excel
	 * inFilePath: 源文件路径
	 * params：替换参数，鉴于参数数量较少，可以将所有参数都放到一个Map中。
	 */
	public String exportExcel(String inFilePath, Map<Integer, Map<String, String>> params){
		String outFilePath = "";
		String outFilePathTemp[] = new String[2];
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inFilePath));
			//TODO 待确认是否写死sheet
			HSSFSheet sheet = workbook.getSheet("居民");
			HSSFRow row = null;
			HSSFCell cell = null;
			
			//获取导出路径
			outFilePathTemp = getOutFilePath(inFilePath);
			outFilePath = outFilePathTemp[0] + outFilePathTemp[1];
			FileOutputStream outStream = new FileOutputStream(outFilePath);
			//循环读取源文件，替换参数
			for(int i = 1; i<= sheet.getLastRowNum(); i++){
				row = sheet.getRow(i);
				for(int columnIndex : params.keySet()){
					cell = row.getCell(columnIndex);
					Map<String, String> parameter = params.get(columnIndex);
					if(cell!=null){
						if(parameter.containsKey(cell.getStringCellValue())){
							cell.setCellValue(parameter.get(cell.getStringCellValue()));
						}
					}
				}
			}
			
			workbook.write(outStream);  
			outStream.close();  	
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return outFilePath;
	}
	
	/*
	 * 导出sql
	 * inFilePath： 预处理后的excel文件路径
	 */
	public void exportSQLFile(String inFilePath){
		
	}
	
	/*
	 * 转换内容，导出excel
	 * inFilePath: 源文件路径
	 * params：替换参数，使用不确定参数列表，可以在执行时手动添加。Integer表示字段列号， Map<String, String>表示参数值及显示名称
	 */
	/*public void exportExcel(String inFilePath, Map<Integer, Map<String, String>> ... params){
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inFilePath));
			//TODO 待确认是否写死sheet
			HSSFSheet sheet = workbook.getSheet("");
			HSSFRow row = null;
			HSSFCell cell = null;
			
			//获取导出路径
			String outFilePath = getOutFilePath(inFilePath);
			FileOutputStream outStream = new FileOutputStream(outFilePath);
			//循环读取源文件，替换参数; POI只能按行读取，无法按列读取，所以需要3次循环，后期有待提高处理速度
			for(int i = 1; i<= sheet.getLastRowNum(); i++){
				row = sheet.getRow(i);
				for(Map<Integer, Map<String, String>> param : params){
					for(int columnIndex : param.keySet()){
						cell = row.getCell(columnIndex);
						Map<String, String> parameter =param.get(columnIndex);
						if(cell!=null){
							if(parameter.containsKey(cell.getStringCellValue())){
								cell.setCellValue(parameter.get(cell.getStringCellValue()));
							}
						}
					}
					
				}
				
			}
			
			
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}*/
	
	
	
	public void replaceParameter(String inFileName, String sheetName, String outFileName, Map<Integer,Map<String, String>> parameters,String endarea) throws IOException{
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inFileName));
		HSSFSheet sheet = workbook.getSheet(sheetName);
		HSSFRow row = null;
		HSSFCell cell = null;
		HSSFCell cell1 = null;
		//Parameter parameter = new Parameter();
		FileOutputStream out = new FileOutputStream(outFileName);  
		for(int i = 1; i<= sheet.getLastRowNum(); i++){
			row = sheet.getRow(i);
			for(int columnIndex:parameters.keySet()){
				//System.out.println("sheetName:"+sheetName+"|i:"+i+"|columnIndex:"+columnIndex);
				cell = row.getCell(columnIndex);
				Map<String, String> parameter =parameters.get(columnIndex);
				if(cell!=null){
					if(parameter.containsKey(cell.getStringCellValue())){
						cell.setCellValue(parameter.get(cell.getStringCellValue()));
					}
				}
			}
			
		}
		if (sheetName.equals("家庭")) {
			
			String zeroStr = "000000000000";

			
			for(int i = 1; i<= sheet.getLastRowNum(); i++){
				row = sheet.getRow(i);
				cell = row.getCell(5);//详细住址
				cell1 = row.createCell(25);//endarea网格地址
				String endareaTemp = "";
				if (null!=cell) {
					switch (cell.getCellType()) {     
                    case HSSFCell.CELL_TYPE_NUMERIC: // 数字                            
                        DecimalFormat df = new DecimalFormat("0");       
                        endareaTemp = df.format(cell.getNumericCellValue());  
                        break;     
                    case HSSFCell.CELL_TYPE_STRING: // 字符串     
                    	endareaTemp = cell.getStringCellValue(); 
                        break;     
					
					}
				}
				endareaTemp = endareaTemp.replaceAll("[\u4e00-\u9fa5]+", "").replace("-", "").replace(" ", "");//去掉汉字
				
				endareaTemp = zeroStr.substring(0,(12-endareaTemp.length()))+endareaTemp;
				
				cell1.setCellValue(endarea+"-"+endareaTemp);
			}
		}	
		workbook.write(out);  
		out.close();  	
	}
	
	/**
	 * 
	 * @param inFileName 导入数据的文件路径 
	 * @param sheetName	工作薄名称
	 * @param outFileName	导出数据的文件路径
	 * @param colunmSum	列数
	 * @param header	头语句
	 * @param areas		城市、县区、街道、社区的编码
	 * @throws IOException
	 */
	public void createFile(String inFileName, String sheetName, String outFileName, int colunmSum, String header, String[] areas) throws IOException{
		//新建一个文件（没有的话，新建一个）
		File file = new File(outFileName);
		file.createNewFile();
		FileWriter fr = new FileWriter(file);
		if(!file.exists())file.createNewFile();
		//使用StringBuffer效率更高，加入头语句
		StringBuffer strb = new StringBuffer();
		strb.append(header);
		
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inFileName));
		HSSFSheet sheet = workbook.getSheet(sheetName);
		HSSFRow row = null;
		HSSFCell cell = null;
		int y = sheet.getLastRowNum();
		int x = y;
		for(int i = 1; i<= sheet.getLastRowNum(); i++){
			row = sheet.getRow(i);
			strb.append("\n");
			strb.append("(");
			strb.append("'"+OperationNoUtil.getUUID()+"',");
			for(int j = 1; j<colunmSum;j++){//过滤数据库里没有字段
				
				if(!(
					(sheetName.equals("家庭")&&j==13)||
					(sheetName.equals("居民")&&(j==18||j==20||j==21||j==22||j==23||j==24||j==25||j==40||j==41||j==45))
					)){
					try{
						cell = row.getCell(j);
						String cellvalue = "";
						if(cell == null){
							strb.append("null,");
							cellvalue = "空值";
						}else{
							//把数字或可能被认为成数字的转换为字符串
							DecimalFormat df = new DecimalFormat("0");
							switch(cell.getCellType()){
							case HSSFCell.CELL_TYPE_NUMERIC:
								strb.append("'" + df.format(cell.getNumericCellValue()) + "',");
								cellvalue = df.format(cell.getNumericCellValue());
								break;
							default:
								strb.append("'" + cell.getStringCellValue() + "',");
								cellvalue = cell.getStringCellValue();
								break;
							}
						}
						//if(sheetName.equals("家庭")){System.out.println(sheetName + " i:" + i + " J:"+ j +" value:"+cellvalue );}
						
					}catch (Exception e) {
						System.out.println(sheetName + " i:" + i + " J:" + j );
					}
				}
			}
			strb.append("'0',");  //state = 0 才是正常数据？？
			//加入 城市、县区、街道、社区的编码
			strb.append(areas[0]);
			strb.append(areas[1]);
			strb.append(areas[2]);
			strb.append(areas[3]);
			
			//去掉最后的逗号
			strb.delete(strb.length()-1,strb.length());
			strb.append("),");
		}
		//去掉最后的逗号，加上分好
		strb.delete(strb.length()-1,strb.length());
		strb.append(";");
		
		fr.append(strb);
		fr.close();
	}
	
	public static void main(String[] args) {
		ConvertExcel ce = new ConvertExcel();
		ParameterController pc = new ParameterController();
		pc.getMapFromDBByCategory("education_degree");
		Map<Integer, Map<String, String>> params = new HashMap<Integer, Map<String, String>>();
		params.put(8, pc.getMapFromDBByCategory("relationship_to_householder"));
		ce.exportExcel("E:/ExcelOut/龙家圈20150624.xls", params);
	}
}
