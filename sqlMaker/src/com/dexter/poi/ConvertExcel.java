package com.dexter.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ConvertExcel {

	public void convertExcel(String filePath){
		String[] paths = filePath.split("/");
		String tempPath = "";	//ת�����excel�ļ�·������·��
		String tempFileName = "";	//ת�����excel����
		for(int i = 0; i < (paths.length - 1); i++){
			tempPath += paths[i] + "/";
		}
		//ת�����excel����Ϊ ������ʱ����_ԭ����
		SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmmss");
		tempFileName += sdf.format(new Date()) + "_" + paths[paths.length-1];
		
		tempPath += tempFileName;
		System.out.println(tempPath);
		
	}
	
	
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
		if (sheetName.equals("��ͥ")) {
			
			String zeroStr = "000000000000";

			
			for(int i = 1; i<= sheet.getLastRowNum(); i++){
				row = sheet.getRow(i);
				cell = row.getCell(5);//��ϸסַ
				cell1 = row.createCell(25);//endarea�����ַ
				String endareaTemp = "";
				if (null!=cell) {
					switch (cell.getCellType()) {     
                    case HSSFCell.CELL_TYPE_NUMERIC: // ����                            
                        DecimalFormat df = new DecimalFormat("0");       
                        endareaTemp = df.format(cell.getNumericCellValue());  
                        break;     
                    case HSSFCell.CELL_TYPE_STRING: // �ַ���     
                    	endareaTemp = cell.getStringCellValue(); 
                        break;     
					
					}
				}
				endareaTemp = endareaTemp.replaceAll("[\u4e00-\u9fa5]+", "").replace("-", "").replace(" ", "");//ȥ������
				
				endareaTemp = zeroStr.substring(0,(12-endareaTemp.length()))+endareaTemp;
				
				cell1.setCellValue(endarea+"-"+endareaTemp);
				
				
				
				
			}
		}	
		workbook.write(out);  
		out.close();  	
	}
	
	
	public static void main(String[] args) {
		ConvertExcel ce = new ConvertExcel();
		ce.convertExcel("D:/tet/test.xls");
	}
}
