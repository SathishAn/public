package utillities_UAT;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;

import org.apache.commons.exec.launcher.Java13CommandLauncher;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadObject {
	static HashMap<String, String> newObject;
	static HashMap<String, String> newattributes;
	static HashMap<String, String> newObjectName;
	static HashMap<String, String> oldObject;
	static HashMap<String, String> oldObjectComments;
	public static void main(String[] args) {
		String path = "./src/test/resources/Datatable/accelator.xlsx";
		newObjectName = new HashMap<String, String>();
		newObject = new HashMap<String, String>();
		newattributes = new HashMap<String, String>();
		File file = new File(path);
		FileInputStream inputStream;
		try {
			inputStream = new FileInputStream(file);
		
		XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getLastRowNum();
		System.out.println(rowCount);
		Row row = sheet.getRow(0);
		String objectName = null, attribute = null, fieldType = null;
		int colCount = row.getLastCellNum();
		for (int i=1 ; i <= rowCount ; i++){
			String fieldText = null;
			row = sheet.getRow(i);
			System.out.println(row.getCell(1).getStringCellValue());
			if (!(row.getCell(0) == null || row.getCell(0).toString().isEmpty() )){
				
				fieldText	= row.getCell(0).getStringCellValue();
			}
			if (!(row.getCell(1) == null)){
				fieldType = row.getCell(1).getStringCellValue();
			}	
			if (!(row.getCell(2) == null || row.getCell(2).toString().isEmpty() )){
				attribute = "ID";
				 objectName	= row.getCell(2).getStringCellValue();
			}else if (!(row.getCell(3) == null || row.getCell(3).toString().isEmpty())){
				attribute = "name";
				 objectName	= row.getCell(3).getStringCellValue();
			}else if (!(row.getCell(5) == null || row.getCell(5).toString().isEmpty())){
				attribute = "Class";
				 objectName	= row.getCell(5).getStringCellValue();
			}else if (!(row.getCell(9) == null || row.getCell(9).toString().isEmpty())){
					attribute = "Link";
				 objectName	= row.getCell(9).getStringCellValue();
			}else if(!(row.getCell(10) == null || row.getCell(10).toString().isEmpty())){
				attribute = "Xpath";
			 objectName	= row.getCell(10).getStringCellValue();
		}
			newObject.put(objectName, fieldType);
			newattributes.put(objectName, attribute);
			if (!(newObjectName.containsKey(objectName))){
				if  (newObjectName.containsValue(fieldText)){
					fieldText= fieldText+"_" + objectName;
					newObjectName.put(objectName, fieldText);
				}else{
					
					newObjectName.put(objectName, fieldText);
				}
			}
			
		}
		
		for (Entry<String, String> entry: newObject.entrySet()){
			System.out.println(entry.getKey() + "-" +entry.getValue());
		}
		for (Entry<String, String> entry: newattributes.entrySet()){
			System.out.println(entry.getKey() + "-" +entry.getValue());
		}
		
		changeClass();
		mapComparison();
		
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	
	public static void changeClass(){
		String className = "ObjectRepository";
		oldObject = new HashMap<String, String>();
		oldObjectComments = new HashMap<String, String>();
		String indexValues = null;
		File javaFile = new File("./src/test/java/utillities_UAT/" +className + ".java");
		try {
			BufferedReader br = new BufferedReader(new FileReader(javaFile));

            String st;
            while((st = br.readLine()) != null){
            	String sPrefix = "public static final By";
            	if (st.contains(sPrefix)){
            		 
            		 int len = sPrefix.length();
            		 String variable = st.substring(st.indexOf(sPrefix));
            		 variable = variable.substring(len);
            				 System.out.println(variable); 
            		String variableKey = variable.substring(1, variable.indexOf("=")).trim();
            		String variableValue = variable.substring(variable.indexOf("=") +1).trim();
            		
            		oldObject.put(variableKey.trim(), variableValue.trim());
            		oldObjectComments.put(variableKey.trim(), indexValues);
            		
            				 
            	}else{
            		if(!st.isEmpty()){
            			indexValues = st;
            		}
            	}
            
            }
            
            for (Entry<String, String> entry: oldObject.entrySet()){
    			System.out.println(entry.getKey() + "-" +entry.getValue());
    		}
		
		
	}catch(Exception e){
		System.out.println(e);
		
	}
		
	}


public static void mapComparison() throws IOException{
	String comparisonValue = null;
	String className = "ObjectRepository";
	ArrayList<String> newlist = new ArrayList<String>();
	HashMap<String, String> propertyChange = new HashMap<String, String>();	
	
	File javaFile = new File("./src/test/java/utillities_UAT/" +className + ".java");
	
	for (String k:newObject.keySet() ){
		
		if(!(newObjectName.get(k)== null || newObjectName.get(k).isEmpty())){
			if (!oldObject.containsKey(newObjectName.get(k))){
				System.out.println(k);
				newlist.add(k);
			}else{
					
					switch(newattributes.get(k)){
					
					case "ID":
						comparisonValue = "By.id(\""+ k +"\");";
						break;
						
					case "name":
						comparisonValue = "By.name(\""+ k +"\");";
						break;
						
					case "Class":
						comparisonValue = "By.className(\""+ k +"\");";
						break;
					case "Link":
						comparisonValue = "By.linkText(\""+ k +"\");";
						break;
					case "Xpath":
						comparisonValue = "By.xpath(\""+ k +"\");";
						break;
					
					default:
						break;
					
					}
								
					System.out.println(oldObject.get(newObjectName.get(k)) + " ------- "+ comparisonValue.toString() );
					
				 if(!oldObject.get(newObjectName.get(k)).equalsIgnoreCase(comparisonValue.toString())){
					 oldObject.put(newObjectName.get(k), comparisonValue);
					 propertyChange.put(newObjectName.get(k), "Yes");
					 
				 }
			}
			
		}else{
			if (!oldObject.containsKey(k)){
				System.out.println(k);
				newlist.add(k);
			}else{
				
				switch(newattributes.get(k)){
				
				case "ID":
					comparisonValue = "By.id(\""+ k +"\");";
					break;
					
				case "name":
					comparisonValue = "By.name(\""+ k +"\");";
					break;
					
				case "Class":
					comparisonValue = "By.className(\""+ k +"\");";
					break;
				case "Link":
					comparisonValue = "By.linkText(\""+ k +"\");";
					break;
				case "Xpath":
					comparisonValue = "By.xpath(\""+ k +"\");";
					break;
				default:
					break;
				
				}
				
				System.out.println(oldObject.get(k) + " ------- "+ comparisonValue.toString() );
				if(!oldObject.get(k).equalsIgnoreCase(comparisonValue.toString())){
					 oldObject.put(k, comparisonValue);
					 propertyChange.put(k, "Yes");
					 
				 }
				
			}
		
		}
	}
	
	BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(javaFile)));
	String FileContent = null;
	String outputValue = null;	
	writer.write("package utillities_UAT;\n\n");
	writer.write("import org.openqa.selenium.By;\n\n");
	writer.write("public class " + className + "{\n\n");	
	writer.write("//Existing Objects");
	for (String key:oldObject.keySet() ){
		if(propertyChange.containsKey(key)){
			String timeStamp = new SimpleDateFormat("dd.MM.yyyy - HH.mm.ss").format(new Date());
			writer.write("\n //Modification done on object property - " + timeStamp);
			
		}else if (!(oldObjectComments.get(key)==null)){
			writer.write("\n\n" + oldObjectComments.get(key) );
		}
		writer.write("\n\n public static final By  " + key + " = " + oldObject.get(key) );
		
	}
	
	
	
	
	
	for(int i=0;i< newlist.size(); i++){
		String newValue;
		if(!(newObjectName.get(newlist.get(i)) == null || newObjectName.get(newlist.get(i)).isEmpty())){
			newValue = newObjectName.get(newlist.get(i));
		}
		else{
			newValue = newlist.get(i);
		}
		String timeStamp = new SimpleDateFormat("dd.MM.yyyy - HH.mm.ss").format(new Date());
		writer.write("\n //New Object added - " + timeStamp);
	writer.write("\n\n public static final By " + newValue  + " = ");
	String sAttribute = newattributes.get(newlist.get(i));
	
	switch (sAttribute) {
	case "ID":
		outputValue = "By.id(\""+ newlist.get(i)+"\")";
		break;
		
	case "name":
		outputValue = "By.name(\""+ newlist.get(i)+"\")";
		break;
		
	case "Class":
		outputValue = "By.className(\""+ newlist.get(i)+"\")";
		break;
	case "Link":
		outputValue = "By.linkText(\""+ newlist.get(i)+"\")";
		break;
	case "Xpath":
		outputValue = "By.xpath(\""+  newlist.get(i) +"\");";
		break;
	
	default:
		break;
	}
	writer.write(outputValue + ";\n");
	
	}
    writer.write("\n}\n");
    writer.close();
    
    
    
    String skelClassName = "SkeletonCode";
    
    File javaskelFile = new File("./src/test/java/utillities_UAT/" +skelClassName + ".java");
    BufferedWriter SkelWriter = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(javaskelFile)));
	
    SkelWriter.write("package utillities_UAT;\n\n");
    SkelWriter.write("import org.openqa.selenium.By;\n\n");
    SkelWriter.write("public class " + skelClassName + "{\n\n");	
    SkelWriter.write("//New Skel Code for newly identifed Objects");
	for(int i=0;i< newlist.size(); i++){
		String new_Value = newObject.get(newlist.get(i));
		String newValue;
		if(!(newObjectName.get(newlist.get(i)) == null || newObjectName.get(newlist.get(i)).isEmpty())){
			newValue = newObjectName.get(newlist.get(i));
		}
		else{
			newValue = newlist.get(i);
		}
		System.out.println(newValue);
		switch (new_Value) {
		case "EditBox":
			SkelWriter.write("//Skeleton code for new object and need to add the feild ' " + newValue+ "' in test data sheet \n\n" );
			SkelWriter.write("enterInputText(\"" + newValue + "\", testdata.get(\"" + newValue + "\"));\n");
			SkelWriter.write("driver.findElement(\""+ newValue +"\").getText();\n\n");
			break;
		case "RadioButton":
			SkelWriter.write("//Skeleton code for new object" + newValue + "\n\n" );
			SkelWriter.write("actionClick(\"" + newValue + "\");\n");	
			break;
			
		case "CheckBox":
			SkelWriter.write("//Skeleton code for new object" + newValue + "\n\n" );
			SkelWriter.write("actionClick(\"" + newValue + "\");\n");	
			break;	
		
			
		case "Button":
			SkelWriter.write("//Skeleton code for new object" + newValue + "\n\n" );
			SkelWriter.write("actionClick(\"" + newValue + "\");\n");	
			SkelWriter.write("driver.findElement(\""+ newValue +"\").getText();\n\n");
			break;
			
		case "Dropdown":
			SkelWriter.write("//Skeleton code for new object and need to add the feild ' " + newValue+ "' in test data sheet \n\n" );
			SkelWriter.write("selectDropdown(\"" + newValue + "\", testdata.get(\"" + newValue + "\"));\n");
			SkelWriter.write("actionClick(\"" + newValue + "\");\n");	
			
			break;
			
		case "Link":
			SkelWriter.write("Skeleton code for new object " + newValue + "\n\n" );
			SkelWriter.write("driver.findElement(\""+newValue +"\").getText();\n\n");
			SkelWriter.write("actionClick(\"" + newValue + "\");\n");
			break;

		default:
			break;		
		
		}
		
		
	}
	SkelWriter.write("\n}\n");
	SkelWriter.close();
	
	
    
}




}
