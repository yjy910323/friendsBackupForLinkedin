package unti;
import java.util.Iterator;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.Region;


public class XLSUnti {
	private String outpath;
	
	public XLSUnti(String outpath){
		this.outpath = outpath;
	}
	
	public static void insertOneData(HSSFRow row,String[] data,HSSFCellStyle s1,HSSFCellStyle s2){
		
		for (int i = 0; i < data.length; i++) {
			HSSFCell cell0000 = row.createCell(i);
			cell0000.setCellValue(data[i]);
			if(i==8||i==21||i==34||i==47||i==52||i==57||i==62){
				cell0000.setCellStyle(s1);
			}
			if(i==7||i==20||i==33||i==46||i==51||i==56||i==61){
				cell0000.setCellStyle(s2);
			}	
		}	
	}
	
	
	public static String loadData(JSONObject result,HSSFSheet sheet,int iid){
	    String[] s = new String[66];
	    s[0] = String.valueOf(iid);
	    if(result.get("firstName")!=null)s[1]=(String) result.get("firstName");
	    if(result.get("lastName")!=null)s[2]=(String) result.get("lastName");
	    if(result.get("headline")!=null)s[3]=(String) result.get("headline");
	    if(result.get("industry")!=null)s[4]=(String) result.get("industry");
	    if(result.get("numConnections")!=null)s[5]=String.valueOf(result.get("numConnections"));
	    if(result.get("pictureUrl")!=null)s[6]=(String) result.get("pictureUrl");
	    if(result.get("publicProfileUrl")!=null)s[7]=(String) result.get("publicProfileUrl");

	    JSONArray positionresult = null;
	    if(result.get("positions")!=null&&JSONObject.fromObject(result.get("positions")).get("values")!=null){
	    	positionresult = JSONObject.fromObject(result.get("positions")).getJSONArray("values");
	    	Iterator<JSONObject> ip = positionresult.iterator();
	    	int i=0;
	    	while(ip.hasNext()){
	    		JSONObject pr = ip.next();
	    		if(pr.get("company")!=null){
	    			JSONObject companyresult = null;
	    			companyresult = (JSONObject) pr.get("company");
	    			if(companyresult.get("id")!=null)s[16+i*13]=String.valueOf(companyresult.get("id"));
	    			if(companyresult.get("industry")!=null)s[19+i*13]=String.valueOf(companyresult.get("industry"));
	    			if(companyresult.get("name")!=null)s[17+i*13]=String.valueOf(companyresult.get("name"));
	    			if(companyresult.get("size")!=null)s[20+i*13]=String.valueOf(companyresult.get("size"));
	    			if(companyresult.get("type")!=null)s[18+i*13]=String.valueOf(companyresult.get("type"));
	    		}
	    		if(pr.get("startDate")!=null){
	    			JSONObject sresult = null;
	    			sresult = (JSONObject) pr.get("startDate");
	    			if(sresult.get("year")!=null)s[11+i*13]=String.valueOf(sresult.get("year"));
	    			if(sresult.get("month")!=null)s[12+i*13]=String.valueOf(sresult.get("month"));
	    		}
	    		if(pr.get("endDate")!=null){
	    			JSONObject eresult = null;
	    			eresult = (JSONObject) pr.get("endDate");
	    			if(eresult.get("year")!=null)s[14+i*13]=String.valueOf(eresult.get("year"));
	    			if(eresult.get("month")!=null)s[13+i*13]=String.valueOf(eresult.get("month"));
	    		}
	    		if(pr.get("isCurrent")!=null)s[15+i*13]=String.valueOf(pr.get("isCurrent"));
	    		if(pr.get("summary")!=null)s[10+i*13]=(String) pr.get("summary");
	    		if(pr.get("title")!=null)s[9+i*13]=(String) pr.get("title");
	    		if(pr.get("id")!=null)s[8+i*13]=String.valueOf(pr.get("id"));
	        			
	    		if(i==3){
	    			break;
	    		}else{
	    			i++;
	    		}
	        			
	    	}
	    }
	    	   
	    	   
	    	   
	    JSONArray educationresult = null;
	    if(result.get("educations")!=null&&JSONObject.fromObject(result.get("educations")).get("values")!=null){
	    	educationresult = JSONObject.fromObject(result.get("educations")).getJSONArray("values");
	    	Iterator<JSONObject> ie = educationresult.iterator();
	        		
	    	int i=0;
	    	while(ie.hasNext()){
	    		JSONObject er = ie.next();
	    		if(er.get("degree")!=null)s[47+i*5]=String.valueOf(er.get("degree"));
	    		if(er.get("activities")!=null)
	    			if(er.get("fieldOfStudy")!=null)s[50+i*5]=String.valueOf(er.get("fieldOfStudy"));
	    		if(er.get("notes")!=null)
	    			if(er.get("schoolName")!=null)s[51+i*5]=String.valueOf(er.get("schoolName"));
	    		if(er.get("startDate")!=null){
	    			JSONObject sresult = null;
	    			sresult = (JSONObject) er.get("startDate");
	    			if(sresult.get("year")!=null)s[48+i*5]=String.valueOf(sresult.get("year"));
	    		}
	    		if(er.get("endDate")!=null){
	    			JSONObject eresult = null;
	    			eresult = (JSONObject) er.get("endDate");
	    			if(eresult.get("year")!=null)s[49+i*5]=String.valueOf(eresult.get("year"));
	    		}
	    		if(i==3){
	    			break;
	    		}else{
	    			i++;
	    		}
	    	}
	        		
	    }
	    	   
	        		
	    	   	
	        		
	        		
	    JSONArray phoneresult = null;
	    if(result.get("phoneNumbers")!=null&&JSONObject.fromObject(result.get("phoneNumbers")).get("values")!=null){
	    	phoneresult = JSONObject.fromObject(result.get("phoneNumbers")).getJSONArray("values");
	    	Iterator<JSONObject> ip = phoneresult.iterator();
	        		
	    	while(ip.hasNext()){
	    		JSONObject pr = ip.next();
	    		if(pr.get("phoneType")!=null)
	    			if(pr.get("phoneNumber")!=null)s[62]=String.valueOf(pr.get("phoneNumber"));
	    	}
	    }
	    	   
	    	   
	    JSONArray skillresult = null;
	    if(result.get("skills")!=null&&JSONObject.fromObject(result.get("skills")).get("values")!=null){
	    	skillresult = JSONObject.fromObject(result.get("skills")).getJSONArray("values");
	    	Iterator<JSONObject> is = skillresult.iterator();
	    	s[63]="";
	    	while(is.hasNext()){
	    		JSONObject sr = is.next();
	    		if(sr.get("skill")!=null){
	    			s[63] += JSONObject.fromObject(sr.get("skill")).get("name") + " | ";
	    		}
	    	}
	    }
	    	   
	        	
	    JSONObject locationresult = null;
	    if(result.get("location")!=null&&JSONObject.fromObject(result.get("location")).get("name")!=null){
	    	locationresult = JSONObject.fromObject(result.get("location"));
	    	s[64] = (String) locationresult.get("name");
	    }
	        	
	        	
	        	
	    JSONArray memberUrlResourcesresult = null;
	    if(result.get("memberUrlResources")!=null&&JSONObject.fromObject(result.get("memberUrlResources")).get("values")!=null){
	    	memberUrlResourcesresult = JSONObject.fromObject(result.get("memberUrlResources")).getJSONArray("values");
	    	Iterator<JSONObject> is = memberUrlResourcesresult.iterator();
	    	while(is.hasNext()){
	    		JSONObject sr = is.next();
	        			
	    		if(sr.get("url")!=null){
	    			s[65] = (String) sr.get("url");
	    		}
	    	}
	    }
	        	
	    	   
	    sheet.setColumnWidth(0,1000);
	    sheet.setColumnWidth(3,8000);
	    sheet.setColumnWidth(4,5000);
	    sheet.setColumnWidth(5,4000);
	    sheet.setColumnWidth(6,7000);
	    sheet.setColumnWidth(7,7000);
	        	
	    for (int i = 0; i < 3; i++) {
	    	sheet.setColumnWidth(8+i*13, 4000);
	    	sheet.setColumnWidth(9+i*13, 8500);
	    	sheet.setColumnWidth(10+i*13, 6500);
	    	sheet.setColumnWidth(11+i*13, 2000);
	    	sheet.setColumnWidth(12+i*13, 2500);
	    	sheet.setColumnWidth(13+i*13, 2000);
	    	sheet.setColumnWidth(14+i*13, 2000);
	    	sheet.setColumnWidth(15+i*13, 2000);
	    	sheet.setColumnWidth(16+i*13, 3000);
	    	sheet.setColumnWidth(17+i*13, 7500);
	    	sheet.setColumnWidth(18+i*13, 5000);
	    	sheet.setColumnWidth(19+i*13, 5500);
	    	sheet.setColumnWidth(20+i*13, 5000);
	    }
	        	
	    for (int i = 0; i < 3; i++) {
	    	sheet.setColumnWidth(47+i*5, 6500);
	    	sheet.setColumnWidth(48+i*5, 2000);
	    	sheet.setColumnWidth(49+i*5, 2000);
	    	sheet.setColumnWidth(50+i*5, 5000);
	    	sheet.setColumnWidth(51+i*5, 10000);
	    }
	    sheet.setColumnWidth(62,4000);
	    sheet.setColumnWidth(63,5000);
	    sheet.setColumnWidth(64,2000);
	    sheet.setColumnWidth(65,8000);
	        	
	    HSSFRow row = sheet.createRow(++iid);	
	}
	
	
	
	public static void createXLSTitle(HSSFSheet sheet,HSSFCellStyle s1,HSSFCellStyle s2){
		try{
			HSSFRow row1 = sheet.createRow(0);
			HSSFCell cell0 = row1.createCell(0);cell0.setCellValue("Basic information");
			Region region0 = new Region((short)0,(short)0,(short)0,(short)7);
			sheet.addMergedRegion(region0); 
			HSSFCell cell1 = row1.createCell(8);cell1.setCellValue("Position 1");
			Region region1 = new Region((short)0,(short)8,(short)0,(short)20);
			sheet.addMergedRegion(region1); 
			HSSFCell cell2 = row1.createCell(21);cell2.setCellValue("Position 2");
			Region region2 = new Region((short)0,(short)21,(short)0,(short)33);
			sheet.addMergedRegion(region2); 
			HSSFCell cell3 = row1.createCell(34);cell3.setCellValue("Position 3");
			Region region3 = new Region((short)0,(short)34,(short)0,(short)46);
			sheet.addMergedRegion(region3); 
			
			HSSFCell cell4 = row1.createCell(47);cell4.setCellValue("Education 1");
			Region region4 = new Region((short)0,(short)47,(short)0,(short)51);
			sheet.addMergedRegion(region4);
			HSSFCell cell5 = row1.createCell(52);cell5.setCellValue("Education 2");
			Region region5 = new Region((short)0,(short)52,(short)0,(short)56);
			sheet.addMergedRegion(region5);
			HSSFCell cell6 = row1.createCell(57);cell6.setCellValue("Education 3");
			Region region6 = new Region((short)0,(short)57,(short)0,(short)61);
			sheet.addMergedRegion(region6);
			
			
			
			
			HSSFCell cell7 = row1.createCell(62);cell7.setCellValue("Contact");
			Region region7 = new Region((short)0,(short)62,(short)0,(short)65);
			sheet.addMergedRegion(region3);
				
			HSSFRow row0 = sheet.createRow(1);
			HSSFCell cell0000 = row0.createCell(0);cell0000.setCellValue("iid");
			HSSFCell cell0001 = row0.createCell(1);cell0001.setCellValue("firstname");
			HSSFCell cell0002 = row0.createCell(2);cell0002.setCellValue("lastname");
			HSSFCell cell0003 = row0.createCell(3);cell0003.setCellValue("headline");
			HSSFCell cell0004 = row0.createCell(4);cell0004.setCellValue("industry");
			HSSFCell cell0005 = row0.createCell(5);cell0005.setCellValue("numconnections");
			HSSFCell cell0006 = row0.createCell(6);cell0006.setCellValue("pictureurl");
			HSSFCell cell0007 = row0.createCell(7);cell0007.setCellValue("publicprofileurl");cell0007.setCellStyle(s2);
			
			HSSFCell cell0008 = row0.createCell(8);cell0008.setCellValue("pid");cell0008.setCellStyle(s1);
			HSSFCell cell0009 = row0.createCell(9);cell0009.setCellValue("title");
			HSSFCell cell0010 = row0.createCell(10);cell0010.setCellValue("summary");
			HSSFCell cell0011 = row0.createCell(11);cell0011.setCellValue("startYear");
			HSSFCell cell0012 = row0.createCell(12);cell0012.setCellValue("startMonth");
			HSSFCell cell0013 = row0.createCell(13);cell0013.setCellValue("endYear");
			HSSFCell cell0014 = row0.createCell(14);cell0014.setCellValue("endMonth");
			HSSFCell cell0015 = row0.createCell(15);cell0015.setCellValue("isCurrent");
			HSSFCell cell0016 = row0.createCell(16);cell0016.setCellValue("companyid");
			HSSFCell cell0017 = row0.createCell(17);cell0017.setCellValue("companyname");
			HSSFCell cell0018 = row0.createCell(18);cell0018.setCellValue("companytype");
			HSSFCell cell0019 = row0.createCell(19);cell0019.setCellValue("companyindustry");
			HSSFCell cell0020 = row0.createCell(20);cell0020.setCellValue("companysize");cell0020.setCellStyle(s2);
			
			HSSFCell cell0021 = row0.createCell(21);cell0021.setCellValue("pid");cell0021.setCellStyle(s1);
			HSSFCell cell0022 = row0.createCell(22);cell0022.setCellValue("title");
			HSSFCell cell0023 = row0.createCell(23);cell0023.setCellValue("summary");
			HSSFCell cell0024 = row0.createCell(24);cell0024.setCellValue("startYear");
			HSSFCell cell0025 = row0.createCell(25);cell0025.setCellValue("startMonth");
			HSSFCell cell0026 = row0.createCell(26);cell0026.setCellValue("endYear");
			HSSFCell cell0027 = row0.createCell(27);cell0027.setCellValue("endMonth");
			HSSFCell cell0028 = row0.createCell(28);cell0028.setCellValue("isCurrent");
			HSSFCell cell0029 = row0.createCell(29);cell0029.setCellValue("companyid");
			HSSFCell cell0030 = row0.createCell(30);cell0030.setCellValue("companyname");
			HSSFCell cell0031 = row0.createCell(31);cell0031.setCellValue("companytype");
			HSSFCell cell0032 = row0.createCell(32);cell0032.setCellValue("companyindustry");
			HSSFCell cell0033 = row0.createCell(33);cell0033.setCellValue("companysize");cell0033.setCellStyle(s2);
			
			HSSFCell cell0034 = row0.createCell(34);cell0034.setCellValue("pid");cell0034.setCellStyle(s1);
			HSSFCell cell0035 = row0.createCell(35);cell0035.setCellValue("title");
			HSSFCell cell0036 = row0.createCell(36);cell0036.setCellValue("summary");
			HSSFCell cell0037 = row0.createCell(37);cell0037.setCellValue("startYear");
			HSSFCell cell0038 = row0.createCell(38);cell0038.setCellValue("startMonth");
			HSSFCell cell0039 = row0.createCell(39);cell0039.setCellValue("endYear");
			HSSFCell cell0040 = row0.createCell(40);cell0040.setCellValue("endMonth");
			HSSFCell cell0041 = row0.createCell(41);cell0041.setCellValue("isCurrent");
			HSSFCell cell0042 = row0.createCell(42);cell0042.setCellValue("companyid");
			HSSFCell cell0043 = row0.createCell(43);cell0043.setCellValue("companyname");
			HSSFCell cell0044 = row0.createCell(44);cell0044.setCellValue("companytype");
			HSSFCell cell0045 = row0.createCell(45);cell0045.setCellValue("companyindustry");
			HSSFCell cell0046 = row0.createCell(46);cell0046.setCellValue("companysize");cell0046.setCellStyle(s2);
			
			
			HSSFCell cell0047 = row0.createCell(47);cell0047.setCellValue("degree");cell0047.setCellStyle(s1);
			HSSFCell cell0048 = row0.createCell(48);cell0048.setCellValue("startDate");
			HSSFCell cell0049 = row0.createCell(49);cell0049.setCellValue("endDate");
			HSSFCell cell0050 = row0.createCell(50);cell0050.setCellValue("fieldOfStudy");
			HSSFCell cell0051 = row0.createCell(51);cell0051.setCellValue("schoolName");cell0051.setCellStyle(s2);
			
			HSSFCell cell0052 = row0.createCell(52);cell0052.setCellValue("degree");cell0052.setCellStyle(s1);
			HSSFCell cell0053 = row0.createCell(53);cell0053.setCellValue("startDate");
			HSSFCell cell0054 = row0.createCell(54);cell0054.setCellValue("endDate");
			HSSFCell cell0055 = row0.createCell(55);cell0055.setCellValue("fieldOfStudy");
			HSSFCell cell0056 = row0.createCell(56);cell0056.setCellValue("schoolName");cell0056.setCellStyle(s2);
			
			HSSFCell cell0057 = row0.createCell(57);cell0057.setCellValue("degree");cell0057.setCellStyle(s1);
			HSSFCell cell0058 = row0.createCell(58);cell0058.setCellValue("startDate");
			HSSFCell cell0059 = row0.createCell(59);cell0059.setCellValue("endDate");
			HSSFCell cell0060 = row0.createCell(60);cell0060.setCellValue("fieldOfStudy");
			HSSFCell cell0061 = row0.createCell(61);cell0061.setCellValue("schoolName");cell0061.setCellStyle(s2);
			
			
			
			HSSFCell cell0062 = row0.createCell(62);cell0062.setCellValue("phoneNumber");cell0062.setCellStyle(s1);
			HSSFCell cell0063 = row0.createCell(63);cell0063.setCellValue("skill");
			HSSFCell cell0064 = row0.createCell(64);cell0064.setCellValue("location");
			HSSFCell cell0065 = row0.createCell(65);cell0065.setCellValue("memberUrlResources");

			
			
		}catch(Exception e){
			e.printStackTrace();
		}
	}
	
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
	}

}
