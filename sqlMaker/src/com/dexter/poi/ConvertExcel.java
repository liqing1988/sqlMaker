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
		//��ȡ�����ļ���
		//ת������������excel
	}
	
	/*
	 * ��ȡ�����ļ���·��
	 * filePath��Դ�ļ�·��
	 * outFilePaths �� 0 �ļ�·����1 ���ɵ��ļ����ƣ� 
	 */
	public String[] getOutFilePath(String inFilePath){
		String[] tempPaths = inFilePath.split("/");
		String[] outFilePaths = new String[2];
		String tempPath = "";	//ת�����excel�ļ�·������·��
		String tempFileName = "";	//ת�����excel����
		for(int i = 0; i < (tempPaths.length - 1); i++){
			tempPath += tempPaths[i] + "/";
		}
		//ת�����excel����Ϊ ������ʱ����_ԭ����
		SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmmss");
		tempFileName += sdf.format(new Date()) + "_" + tempPaths[tempPaths.length-1];
		
		outFilePaths[0] = tempPath;
		outFilePaths[1] = tempFileName;
		return outFilePaths;
	}
	
	/*
	 * ��ȡ����SQL�ļ���·��
	 * filePath��Դ�ļ�·��
	 * sheetName : ����������
	 */
	public String getOutFilePath(String inFilePath, String sheetName){
		String[] paths = inFilePath.split("/");
		String tempPath = "";	//ת�����excel�ļ�·������·��
		String tempFileName = "";	//ת�����excel����
		for(int i = 0; i < (paths.length - 1); i++){
			tempPath += paths[i] + "/";
		}
		//ת�����excel����Ϊ ������ʱ����_ԭ����
		SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmmss");
		tempFileName += sdf.format(new Date()) + "_" + paths[paths.length-1];
		
		tempPath += tempFileName;
		return tempPath;
	}
	
	/*
	 * ת�����ݣ�����excel
	 * inFilePath: Դ�ļ�·��
	 * params���滻���������ڲ����������٣����Խ����в������ŵ�һ��Map�С�
	 */
	public String exportExcel(String inFilePath, Map<Integer, Map<String, String>> params){
		String outFilePath = "";
		String outFilePathTemp[] = new String[2];
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inFilePath));
			//TODO ��ȷ���Ƿ�д��sheet
			HSSFSheet sheet = workbook.getSheet("����");
			HSSFRow row = null;
			HSSFCell cell = null;
			
			//��ȡ����·��
			outFilePathTemp = getOutFilePath(inFilePath);
			outFilePath = outFilePathTemp[0] + outFilePathTemp[1];
			FileOutputStream outStream = new FileOutputStream(outFilePath);
			//ѭ����ȡԴ�ļ����滻����
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
	 * ����sql
	 * inFilePath�� Ԥ������excel�ļ�·��
	 */
	public void exportSQLFile(String inFilePath){
		String familyHeader4 = "use communityservicesys;\n"
				+ "insert into cs_baseinfo_family (`uuid`,`grid_code`,`code`,`house_holder_name`,`house_holder_num`,`address`,`family_register_type`,`if_single_parent`,`if_lone_old_man`,"
				+ "`if_subsistence_family`,`subsistence_money`,`subsistence_begindate`,`subsistence_persons`,`if_sw_old_man`,`if_kc_old_man`,`if_disabled_family`,`difficulty_family_type`,`special_care_type`,"
				+ "`overseas_relations_type`,`income_type`,`if_only_child_family`,`if_exceed_bear_family`,`son_num`,`daughter_num`,`end_area`, `state`, `city`, `area`, `street`, `village`) values ";
		String houseHeader4 = "use communityservicesys;\n"	
				+ "insert into cs_baseinfo_house(`uuid`, `code`, `owner`, `contract_num`, `owner_id`, `house_address`, `house_nature`, `property_right`, `structure_type`, `house_type`, "
				+ "`covered_area`, `usable_area`, `ownership_certificate_num`, `land_use_certificate_num`, `if_have_yard`, `yard_area`, `if_have_storeroom`, "
				+ "`storeroom_area`, `if_have_temporary_building`, `temporary_building_area`, `precautionary_measure`, `state`, `city`, `area`, `street`, `village`) values ";
		String inhabitantHeader4 = "use communityservicesys;\n"
				+ "insert into cs_baseinfo_inhabitant(`uuid`, `code_linshi`, `name`, `country`, `nationality`, `gender`, `accounts_nature`, `contact_num`, " +
						"`relationship_to_householder`, `certificate_type`, `certificate_num`, `marital_status`, `marriage_date`,`education_degree`," +
						"`religion_type`,`endowment_insurance_type`,`medical_insurance_type`,`politics_status`,`ifvillagemanager`,`employment_status`," +
						"`work_unit`,`start_work_date`,`underemployed_reason`,`job_intension`,`voluntary_type`,`if_retirement`,`retirement_date`," +
						"`domicile_type`,`domicile_place`,`if_resident`,`if_migration`,`flowtime`,`flow_reasons`,`cancel_reasons`,`if_immigration`," +
						"`immigration_date`,`military_service`,`arm_name`,`if_disabled_person`,`disabled_type`,`disabled_leavel`,`if_sw_old_man`," +
						"`if_kc_old_man`,`if_sn_old_man`,`religion_manager`,`if_deal_foreign`,`if_problem_teenager`,`if_drug_related_person`," +
						"`if_upder_control_person`,`if_force_person`,`if_concern_borderland`,`if_severe_mental_illness`,`if_community_correction`," +
						"`if_llsf_person`,`if_xmsf_person`,`if_custody_person`, `state`, `city`, `area`, `street`, `village`) values ";
		String[] areas = new String[4];
		/*areas[0] = "'2ff27e17-8dc5-4447-aecc-6ca565509ed5',";   	//����
		areas[1] = "'6c24a747-64c4-4717-8939-cad230e5678e',";		//����
		areas[2] = "'3424a7f3-50e2-4503-bd64-beab63d2a2ed',";		//�ֵ�
		areas[3] = "'31d577b7-aba8-49db-a875-c3e7d568f31b',";		//����
		//areas[4] = "";      //¥�ĸ��ڵ��ǰ24λ
		
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\��ݽ�ƴ���Ϣ.xls", "��ͥ", "g:\\minzhengjuexcelimport\\��ݽ�ƴ���Ϣ0122F.sql", 23, familyHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\��ݽ�ƴ���Ϣ.xls", "����", "g:\\minzhengjuexcelimport\\��ݽ�ƴ���Ϣ0122H.sql", 21, houseHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\��ݽ�ƴ���Ϣ.xls", "����", "g:\\minzhengjuexcelimport\\��ݽ�ƴ���Ϣ0122I.sql", 54, inhabitantHeader4, areas);
		

		sqltemplate.createFile("G:\\minzhengjuexcelimport\\�������Ϣ.xls", "��ͥ", "g:\\minzhengjuexcelimport\\�������Ϣ0122F.sql", 23, familyHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\�������Ϣ.xls", "����", "g:\\minzhengjuexcelimport\\�������Ϣ0122H.sql", 21, houseHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\�������Ϣ.xls", "����", "g:\\minzhengjuexcelimport\\�������Ϣ0122I.sql", 54, inhabitantHeader4, areas);


		sqltemplate.createFile("G:\\minzhengjuexcelimport\\ǰ������Ϣ.xls", "��ͥ", "g:\\minzhengjuexcelimport\\ǰ������Ϣ0122F.sql", 23, familyHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\ǰ������Ϣ.xls", "����", "g:\\minzhengjuexcelimport\\ǰ������Ϣ0122H.sql", 21, houseHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\ǰ������Ϣ.xls", "����", "g:\\minzhengjuexcelimport\\ǰ������Ϣ0122I.sql", 54, inhabitantHeader4, areas);
		
		
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\Ȫ�ڴ���Ϣ.xls", "��ͥ", "g:\\minzhengjuexcelimport\\Ȫ�ڴ���Ϣ0122F.sql", 23, familyHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\Ȫ�ڴ���Ϣ.xls", "����", "g:\\minzhengjuexcelimport\\Ȫ�ڴ���Ϣ0122H.sql", 21, houseHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\Ȫ�ڴ���Ϣ.xls", "����", "g:\\minzhengjuexcelimport\\Ȫ�ڴ���Ϣ0122I.sql", 54, inhabitantHeader4, areas);
		
		
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "��ͥ", "g:\\minzhengjuexcelimport\\������0124F.sql", 23, familyHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "����", "g:\\minzhengjuexcelimport\\������0124H.sql", 21, houseHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "����", "g:\\minzhengjuexcelimport\\������0124I.sql", 54, inhabitantHeader4, areas);
		

		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "��ͥ", "g:\\minzhengjuexcelimport\\������0124F.sql", 23, familyHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "����", "g:\\minzhengjuexcelimport\\������0124H.sql", 21, houseHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "����", "g:\\minzhengjuexcelimport\\������0124I.sql", 54, inhabitantHeader4, areas);


		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "��ͥ", "g:\\minzhengjuexcelimport\\������0124F.sql", 23, familyHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "����", "g:\\minzhengjuexcelimport\\������0124H.sql", 21, houseHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "����", "g:\\minzhengjuexcelimport\\������0124I.sql", 54, inhabitantHeader4, areas);
		
		
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "��ͥ", "g:\\minzhengjuexcelimport\\������0124F.sql", 23, familyHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "����", "g:\\minzhengjuexcelimport\\������0124H.sql", 21, houseHeader4, areas);
		sqltemplate.createFile("G:\\minzhengjuexcelimport\\������.xls", "����", "g:\\minzhengjuexcelimport\\������0124I.sql", 54, inhabitantHeader4, areas);
	
		*/
		areas[0] = "'2ff27e17-8dc5-4447-aecc-6ca565509ed5',";   	//����
		areas[1] = "'60c2d520-2579-4ece-9b29-0c42cd698fdd',";		//����
		areas[2] = "'cf82ffd4-1ab6-4ea7-be1f-d9d4737a6f1a',";		//�ֵ�
		areas[3] = "'9f8c948a-c924-4d52-a0c1-097ee29dfc8f',";		//����
		//areas[4] = "";      //¥�ĸ��ڵ��ǰ24λ
		
		try {
			createFile("G:\\minzhengjuexcelimport\\����Ȧ20150624.xls", "��ͥ", "g:\\minzhengjuexcelimport\\����ȦF.sql", 26, familyHeader4, areas);
			createFile("G:\\minzhengjuexcelimport\\����Ȧ20150624.xls", "����", "g:\\minzhengjuexcelimport\\����ȦH.sql", 21, houseHeader4, areas);
			createFile("G:\\minzhengjuexcelimport\\����Ȧ20150624.xls", "����", "g:\\minzhengjuexcelimport\\����ȦI.sql", 66, inhabitantHeader4, areas);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	/*
	 * ת�����ݣ�����excel
	 * inFilePath: Դ�ļ�·��
	 * params���滻������ʹ�ò�ȷ�������б�������ִ��ʱ�ֶ���ӡ�Integer��ʾ�ֶ��кţ� Map<String, String>��ʾ����ֵ����ʾ����
	 */
	/*public void exportExcel(String inFilePath, Map<Integer, Map<String, String>> ... params){
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inFilePath));
			//TODO ��ȷ���Ƿ�д��sheet
			HSSFSheet sheet = workbook.getSheet("");
			HSSFRow row = null;
			HSSFCell cell = null;
			
			//��ȡ����·��
			String outFilePath = getOutFilePath(inFilePath);
			FileOutputStream outStream = new FileOutputStream(outFilePath);
			//ѭ����ȡԴ�ļ����滻����; POIֻ�ܰ��ж�ȡ���޷����ж�ȡ��������Ҫ3��ѭ���������д���ߴ����ٶ�
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
	
	/**
	 * 
	 * @param inFileName �������ݵ��ļ�·�� 
	 * @param sheetName	����������
	 * @param outFileName	�������ݵ��ļ�·��
	 * @param colunmSum	����
	 * @param header	ͷ���
	 * @param areas		���С��������ֵ��������ı���
	 * @throws IOException
	 */
	public void createFile(String inFileName, String sheetName, String outFileName, int colunmSum, String header, String[] areas) throws IOException{
		//�½�һ���ļ���û�еĻ����½�һ����
		File file = new File(outFileName);
		file.createNewFile();
		FileWriter fr = new FileWriter(file);
		if(!file.exists())file.createNewFile();
		//ʹ��StringBufferЧ�ʸ��ߣ�����ͷ���
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
			for(int j = 1; j<colunmSum;j++){//�������ݿ���û���ֶ�
				
				if(!(
					(sheetName.equals("��ͥ")&&j==13)||
					(sheetName.equals("����")&&(j==18||j==20||j==21||j==22||j==23||j==24||j==25||j==40||j==41||j==45))
					)){
					try{
						cell = row.getCell(j);
						String cellvalue = "";
						if(cell == null){
							strb.append("null,");
							cellvalue = "��ֵ";
						}else{
							//�����ֻ���ܱ���Ϊ�����ֵ�ת��Ϊ�ַ���
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
						//if(sheetName.equals("��ͥ")){System.out.println(sheetName + " i:" + i + " J:"+ j +" value:"+cellvalue );}
						
					}catch (Exception e) {
						System.out.println(sheetName + " i:" + i + " J:" + j );
					}
				}
			}
			strb.append("'0',");  //state = 0 �����������ݣ���
			//���� ���С��������ֵ��������ı���
			strb.append(areas[0]);
			strb.append(areas[1]);
			strb.append(areas[2]);
			strb.append(areas[3]);
			
			//ȥ�����Ķ���
			strb.delete(strb.length()-1,strb.length());
			strb.append("),");
		}
		//ȥ�����Ķ��ţ����Ϸֺ�
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
		ce.exportExcel("E:/ExcelOut/����Ȧ20150624.xls", params);
	}
}
