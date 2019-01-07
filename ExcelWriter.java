package utillities_UAT;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.By.ById;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import utillities.BaseTestSetup;

public class ExcelWriter {
	static int newFlag = 0;
	static XSSFWorkbook wwb;
	static Sheet sheet;
	static File filePath;
	
	public static void designAccelator() throws Exception {
		
		 filePath = new File("./src/test/resources/Datatable/accelator.xlsx");
		if (!filePath.exists() || newFlag == 0){
			createWorkbook();
			newFlag = 1;
		}
		headerUpdateWorkbook();
		
		
		java.util.List<org.openqa.selenium.WebElement> linkElements = BaseTestSetup.driver.findElements(By.tagName("input"));
		for (int i = 0; i < linkElements.size(); i++) {
			if(!linkElements.get(i).getAttribute("type").equals("") & linkElements.get(i).getAttribute("type").equals("text")){
				String placeholder = linkElements.get(i).getAttribute("placeholder");				
				String id = linkElements.get(i).getAttribute("id");				
				String name = linkElements.get(i).getAttribute("name");				
				String classname = linkElements.get(i).getAttribute("class");
				String fieldType = "EditBox";
				System.out.println("Entered fetching objects");
				String xpath = xpathGenerator(linkElements.get(i), "text");
				String FieldText = xpath.substring(xpath.indexOf("~$")+2);
				xpath = xpath.substring(0,xpath.indexOf("~$"));
				dataUpdateWorkbook(FieldText,fieldType, id, name, classname, "", "","", "","", xpath);				
				
			}else if(!linkElements.get(i).getAttribute("type").equals("") & linkElements.get(i).getAttribute("type").equals("password")){
				String placeholder = linkElements.get(i).getAttribute("placeholder");				
				String id = linkElements.get(i).getAttribute("id");				
				String name = linkElements.get(i).getAttribute("name");				
				String classname = linkElements.get(i).getAttribute("class");
				String fieldType = "EditBox";
				System.out.println("Entered fetching objects");
				String xpath = xpathGenerator(linkElements.get(i), "password");
				String FieldText = xpath.substring(xpath.indexOf("~$")+2).trim();
				xpath = xpath.substring(0,xpath.indexOf("~$"));
				
					FieldText = id;
				
				dataUpdateWorkbook(FieldText,fieldType, id, name, classname, "", "","", "", "", xpath);	
			}else if(!linkElements.get(i).getAttribute("type").equals("") & linkElements.get(i).getAttribute("type").equals("radio")){
				String placeholder = linkElements.get(i).getAttribute("placeholder");				
				String id = linkElements.get(i).getAttribute("id");				
				String name = linkElements.get(i).getAttribute("name");				
				String classname = linkElements.get(i).getAttribute("class");
				String fieldType = "RadioButton";
				System.out.println("Entered fetching objects");
				String xpath = xpathGenerator(linkElements.get(i), "checkbox");
				String FieldText = xpath.substring(xpath.indexOf("~$")+2);
				xpath = xpath.substring(0,xpath.indexOf("~$"));
				if (FieldText.equals("") || FieldText == null){
					FieldText = id;
				}
				
				dataUpdateWorkbook(name,fieldType, id, name, classname, "", "","", "", "",xpath);	
			}else if(!linkElements.get(i).getAttribute("type").equals("") & linkElements.get(i).getAttribute("type").equals("checkbox")){
				String placeholder = linkElements.get(i).getAttribute("placeholder");				
				String id = linkElements.get(i).getAttribute("id");				
				String name = linkElements.get(i).getAttribute("name");				
				String classname = linkElements.get(i).getAttribute("class");
				String fieldType = "CheckBox";
				System.out.println("Entered fetching objects");
				String xpath = xpathGenerator(linkElements.get(i), "checkbox");
				String FieldText = xpath.substring(xpath.indexOf("~$")+2);
				xpath = xpath.substring(0,xpath.indexOf("~$"));
				if (FieldText.equals("") || FieldText == null){
					FieldText = id;
				}
				
				dataUpdateWorkbook(name,fieldType, id, name, classname, "", "","", "", "",xpath);	
			}else if(!linkElements.get(i).getAttribute("type").equals("") & linkElements.get(i).getAttribute("type").equals("submit")){
				String placeholder = linkElements.get(i).getAttribute("placeholder");				
				String id = linkElements.get(i).getAttribute("id");				
				String name = linkElements.get(i).getAttribute("name");				
				String classname = linkElements.get(i).getAttribute("class");
				String filedvalue = linkElements.get(i).getText();
				String fieldType = "Button";
				System.out.println("Entered fetching objects");
				dataUpdateWorkbook(filedvalue,fieldType, id, name, classname, "", "","", "","","");
			}else if(!linkElements.get(i).getAttribute("type").equals("text") & linkElements.get(i).getAttribute("value").equals("select")){
				String placeholder = linkElements.get(i).getAttribute("placeholder");				
				String id = linkElements.get(i).getAttribute("id");				
				String name = linkElements.get(i).getAttribute("name");				
				String classname = linkElements.get(i).getAttribute("class");
				String filedvalue = linkElements.get(i).getText();
				String fieldType = "Select";
				System.out.println("Entered fetching objects");
				String xpath = xpathGenerator(linkElements.get(i), "select");
				String FieldText = xpath.substring(xpath.indexOf("~$")+2);
				xpath = xpath.substring(0,xpath.indexOf("~$"));				
					FieldText = id;
				
				
				
				dataUpdateWorkbook(id,fieldType, id, name, classname, "", "","", "", "",xpath);
			}
			
			
			
		}
		
		/*// WebElement link
		java.util.List<org.openqa.selenium.WebElement> linkElementsa = BaseTestSetup.driver
				.findElements(By.tagName("a"));
		for (int i = 0; i < linkElementsa.size(); i++) {
			System.out.println(i + "link");
			if(!linkElementsa.get(i).getText().equals("")){
				String text = linkElementsa.get(i).getText();
				String filedvalue = linkElementsa.get(i).getAttribute("questionid");
				if (!text.isEmpty()){
					System.out.println(text + " ----- " + filedvalue);
				}
				System.out.println("Entered fetching objects");
				//dataUpdateWorkbook(filedvalue,"Link", "", "", "", "", "","", "",text);
			}
		}*/

		
		
		
		// WebElement link
				java.util.List<org.openqa.selenium.WebElement> linkElements2 = BaseTestSetup.driver
						.findElements(By.tagName("a"));
				for (int i = 0; i < linkElements2.size(); i++) {
					System.out.println(i + "link");
					if(!linkElements2.get(i).getText().equals("")){
						String text = linkElements2.get(i).getText();
						String filedvalue = linkElements2.get(i).getText();
						System.out.println("Entered fetching objects");
						dataUpdateWorkbook(filedvalue,"Link", "", "", "", "", "","", "",text, "");
					}
				}
		
				//ListBox
				java.util.List<org.openqa.selenium.WebElement> linkElements3 = BaseTestSetup.driver
						.findElements(By.tagName("select"));
				for (int i = 0; i < linkElements3.size(); i++) {
					System.out.println(i + "select");
					if(!linkElements3.get(i).getAttribute("type").equals("")){
						String placeholder = linkElements3.get(i).getAttribute("placeholder");
						String id = linkElements3.get(i).getAttribute("id");
						String name = linkElements3.get(i).getAttribute("name");
						String classname = linkElements3.get(i).getAttribute("class");
						String filedvalue = linkElements3.get(i).getText();
						System.out.println("Entered fetching objects");
						String xpath = xpathGenerator(linkElements3.get(i), "select");
						String FieldText = xpath.substring(xpath.indexOf("~$")+2);
						xpath = xpath.substring(0,xpath.indexOf("~$"));
						
							FieldText = id;
						
						dataUpdateWorkbook(FieldText,"Dropdown", id, name, classname, "", "","", "", "",xpath);
						
					}
		
				}
		
		
		
		
	}
	//Create a new Workbook
	public static void createWorkbook(){
		try {
			System.out.println("Create Workbook");
			FileOutputStream outStream= new FileOutputStream(filePath);		
			wwb=new XSSFWorkbook();
			sheet = wwb.createSheet("PSTC");
			wwb.write(outStream);
			outStream.close();
		} catch (Exception e) {			// TODO Auto-generated catch block
			e.printStackTrace();
		}
					
	}
	
public static void headerUpdateWorkbook() throws IOException{
		FileInputStream inputstreams = new FileInputStream(filePath);
		wwb= new XSSFWorkbook(inputstreams);
		sheet = wwb.getSheetAt(0);				
		int rowCount = sheet.getLastRowNum();
		System.out.println(rowCount);
	
	if (rowCount < 1){
		
		Row row = sheet.createRow(0);
		row.createCell(0).setCellValue("Field_Text");
		row.createCell(1).setCellValue("Field_Type");
		row.createCell(2).setCellValue("Attribute_ID");
		row.createCell(3).setCellValue("Attribute_Name");
		row.createCell(4).setCellValue("Attribute_InnerText");
		row.createCell(5).setCellValue("Attribute_Class");
		row.createCell(6).setCellValue("Attribute_Value");
		row.createCell(7).setCellValue("Attribute_Placeholder");
		row.createCell(8).setCellValue("Attribute_Title");
		row.createCell(9).setCellValue("Attribute_TextValue");		
		row.createCell(10).setCellValue("Attribute_xpath");		
		
	}
	
	inputstreams.close();
	FileOutputStream outStream = new FileOutputStream(filePath);
	wwb.write(outStream);
	outStream.close();
		
	}

public static void dataUpdateWorkbook(String sf1, String sf2, String sf3,String sf4 , String sf5,String sf6, String sf7, String sf8,String sf9, String sf10 , String sf11) throws IOException{
	System.out.println("Sample");
	FileInputStream inputstreams = new FileInputStream(filePath);
	wwb= new XSSFWorkbook(inputstreams);
	sheet = wwb.getSheetAt(0);				
	int rowCount = sheet.getLastRowNum();
	System.out.println(rowCount);

	Row row = sheet.createRow(rowCount+1);
	System.out.println(sf1);
	row.createCell(0).setCellValue(sf1.replaceAll(" ", "_"));
	row.createCell(1).setCellValue(sf2);
	row.createCell(2).setCellValue(sf3);
	row.createCell(3).setCellValue(sf4);	
	row.createCell(4).setCellValue(sf5);
	row.createCell(5).setCellValue(sf6);
	row.createCell(6).setCellValue(sf7);
	row.createCell(7).setCellValue(sf8);
	row.createCell(8).setCellValue(sf9);
	row.createCell(9).setCellValue(sf10);
	row.createCell(10).setCellValue(sf11);
	

inputstreams.close();
FileOutputStream outStream = new FileOutputStream(filePath);
wwb.write(outStream);
outStream.close();
	
}


public static String xpathGenerator(WebElement linkElements, String type) throws InterruptedException {
	
	// TODO Auto-generated method stub
	int identifiedFlag = 0;
	/*System.setProperty("webdriver.chrome.driver", "./drivers/chromedriver.exe");
	WebDriver driver=new ChromeDriver();		
	driver.manage().window().maximize();
	Thread.sleep(2000);
	driver.get("http://store.demoqa.com/products-page/your-account/");
	Thread.sleep(2000);	*/
	String newXpath = null, tagName;
	
	//List<WebElement>  linkElements = driver.findElements(By.tagName("input"));
	//WebElement linkElements=  driver.findElement(By.id("identifierId"));
	tagName = linkElements.getTagName();
	System.out.println("------"+tagName);
	//newXpath = tagName;
	String[] strFlow = {"Following","Preceding", "Parent", "Parent_sibling", "Parent_sib_child", "Grand_Parent", "Grand_Parent_sibling", "Grand_ParentFollow_sibling"};
//	for (int i = 0; i < linkElements.size(); i++) {
		
//		if(!linkElements.get(i).getAttribute("type").equals("") & linkElements.get(i).getAttribute("type").equals("text")){
			tagName = linkElements.getTagName();
			newXpath = tagName+"[@type='" + type + "']";
			identifiedFlag = 0;
			String xPathValue;
			String attribute, label, value = null;
			List<WebElement> preceding,Parent_Sibling,Grand_Parent_Sibling ;
			WebElement Parents = null, Grand_Parents = null ;
			int iCount = 0;
			do{
				System.out.println("###################################################################");
				System.out.println(strFlow[iCount]);
				switch (strFlow[iCount]) {
				
				case "Following":
					if(!(type.equals("select"))){
						preceding= linkElements.findElements(By.xpath("following-sibling::*"));
						System.out.println(preceding.size());
						if (preceding.size() > 0){
							System.out.println("Entering into precedeing");
							xPathValue = findElementsTag(preceding);
							if (xPathValue != null){
								attribute = xPathValue.substring(0, xPathValue.indexOf("~#"));
								System.out.println(attribute);
								label = xPathValue.substring(xPathValue.indexOf("~#")+2, xPathValue.indexOf("~*"));
								System.out.println(label);
								value = xPathValue.substring(xPathValue.indexOf("~*")+2);
								System.out.println(value);
								String xValue = generateXpath(attribute,label, value);
								System.out.println(xValue);
								newXpath =xValue +"/preceding-sibling::" + newXpath;
								identifiedFlag = 1;
								System.out.println(newXpath);
							}
						}
					}
					
					break;
				
				case "Preceding":
					preceding= linkElements.findElements(By.xpath("preceding-sibling::*"));
					System.out.println(preceding.size());
					if (preceding.size() > 0){
						System.out.println("Entering into precedeing");
						xPathValue = findElementsTag(preceding);
						if (xPathValue != null){
							attribute = xPathValue.substring(0, xPathValue.indexOf("~#"));
							System.out.println(attribute);
							label = xPathValue.substring(xPathValue.indexOf("~#")+2, xPathValue.indexOf("~*"));
							System.out.println(label);
							value = xPathValue.substring(xPathValue.indexOf("~*")+2);
							System.out.println(value);
							String xValue = generateXpath(attribute,label, value);
							System.out.println(xValue);
							newXpath =xValue +"/following-sibling::" + newXpath;
							identifiedFlag = 1;
							System.out.println(newXpath);
						}
					}
					break;
				case "Parent":
					Parents = linkElements.findElement(By.xpath("parent::*"));
					tagName = Parents.getTagName();
					System.out.println("Entering into parents " + tagName);
					xPathValue = findParentTag(Parents);
					if (xPathValue != null){
						attribute = xPathValue.substring(0, xPathValue.indexOf("~#"));
						System.out.println(attribute);
						label = xPathValue.substring(xPathValue.indexOf("~#")+2, xPathValue.indexOf("~*"));
						System.out.println(label);
						value = xPathValue.substring(xPathValue.indexOf("~*")+2);
						System.out.println(value);
						String xValue = generateXpath(attribute,label, value);
						System.out.println(xValue);
						newXpath =xValue +"/" + newXpath;
						identifiedFlag = 1;
						System.out.println(newXpath);
					}else{
						newXpath = tagName + "/" + newXpath;
					}
					break;
				case "Parent_sibling":
					Parent_Sibling = Parents.findElements(By.xpath("preceding-sibling::*"));
					if (Parent_Sibling.size() > 0){
						System.out.println("Entering into Parent precedeing");
						xPathValue = findElementsTag(Parent_Sibling);
						System.out.println(xPathValue);
						if (xPathValue != null){
							attribute = xPathValue.substring(0, xPathValue.indexOf("~#"));
							System.out.println(attribute);
							label = xPathValue.substring(xPathValue.indexOf("~#")+2, xPathValue.indexOf("~*"));
							System.out.println(label);
							value = xPathValue.substring(xPathValue.indexOf("~*")+2);
							System.out.println(value);
							String xValue = generateXpath(attribute,label, value);
							System.out.println(xValue);
							newXpath =xValue +"/following-sibling::" + newXpath;
							identifiedFlag = 1;
							System.out.println(newXpath);
						}
					}	
					break;
				case "Grand_Parent":
					Grand_Parents = Parents.findElement(By.xpath("parent::*"));
					tagName = Grand_Parents.getTagName();
					System.out.println("Entering into Grand parents " + tagName);
					xPathValue = findParentTag(Grand_Parents);
					if (xPathValue != null){
						attribute = xPathValue.substring(0, xPathValue.indexOf("~#"));
						System.out.println(attribute);
						label = xPathValue.substring(xPathValue.indexOf("~#")+2, xPathValue.indexOf("~*"));
						System.out.println(label);
						value = xPathValue.substring(xPathValue.indexOf("~*")+2);
						System.out.println(value);
						String xValue = generateXpath(attribute,label, value);
						System.out.println(xValue);
						newXpath =xValue +"/" + newXpath;
						identifiedFlag = 1;
						System.out.println(newXpath);
					}else{
						newXpath = tagName + "/" + newXpath;
					}
					break;
				case "Grand_Parent_sibling":
					Grand_Parent_Sibling = Grand_Parents.findElements(By.xpath("preceding-sibling::*"));
					if (Grand_Parent_Sibling.size() > 0){
						System.out.println("Entering into Grand Parent precedeing");
						xPathValue = findElementsTag(Grand_Parent_Sibling);
						System.out.println(xPathValue);
						if (xPathValue != null){
							attribute = xPathValue.substring(0, xPathValue.indexOf("~#"));
							System.out.println(attribute);
							label = xPathValue.substring(xPathValue.indexOf("~#")+2, xPathValue.indexOf("~*"));
							System.out.println(label);
							value = xPathValue.substring(xPathValue.indexOf("~*")+2);
							System.out.println(value);
							String xValue = generateXpath(attribute,label, value);
							System.out.println(xValue);
							newXpath =xValue +"/following-sibling::" + newXpath;
							identifiedFlag = 1;
							System.out.println(newXpath);
						}else{
							
							List<WebElement> child= Grand_Parent_Sibling.get(0).findElements(By.xpath("descendant::*"));
							System.out.println(child.size());
							if (child.size() > 0){
								System.out.println("Entering into Descendant");
								xPathValue = findElementsTag(child);
								System.out.println(xPathValue);
								if (xPathValue != null){
									newXpath = "//label[contains(text(), '" + xPathValue + "')]/../following-sibling::" + newXpath;
									System.out.println(newXpath);
								}
							}
						}
					}	
					break;
				case "Grand_ParentFollow_sibling":
					Grand_Parent_Sibling = Grand_Parents.findElements(By.xpath("following-sibling::*"));
					if (Grand_Parent_Sibling.size() > 0){
						System.out.println("Entering into Grand Parent Following");
						xPathValue = findElementsTag(Grand_Parent_Sibling);
						System.out.println(xPathValue);
						if (xPathValue != null){
							attribute = xPathValue.substring(0, xPathValue.indexOf("~#"));
							System.out.println(attribute);
							label = xPathValue.substring(xPathValue.indexOf("~#")+2, xPathValue.indexOf("~*"));
							System.out.println(label);
							value = xPathValue.substring(xPathValue.indexOf("~*")+2);
							System.out.println(value);
							String xValue = generateXpath(attribute,label, value);
							System.out.println(xValue);
							newXpath =xValue +"/preceding-sibling::" + newXpath;
							identifiedFlag = 1;
							System.out.println(newXpath);
						}
						
					}
					break;
				default:
					break;
				}
				iCount++;
				
				
			}while(identifiedFlag == 0 & iCount < strFlow.length );
			if(value !=null){
				if (value.length() > 25){
					value = value.substring(0, value.indexOf("\n"));
				}
			}
			
			return newXpath + "~$" + value;
			
		
		
		
		
		


}

public static String findElementsTag(List<WebElement> element){
	String tagName = null;
	System.out.println(element.get(0).getTagName());
	for (int j = 0; j< element.size(); j++){
		System.out.println(element.get(j).getTagName());
		if (element.get(j).getAttribute("for") != null){
			tagName =  "FOR"  + "~#" + element.get(j).getTagName() + "~*" + element.get(j).getAttribute("for") ;
		}else if (!element.get(j).getText().equals("")){
			System.out.println(element.get(j).getTagName());
			List<WebElement> childElements = element.get(j).findElements(By.xpath("descendant::*"));
			if (childElements.size() >0){
				for(int i =0; i < childElements.size();i++){
					if (childElements.get(i).getText() != null && !(childElements.get(i).getText().equals(""))){
						tagName = "Text-decendent"  + "~#" + childElements.get(i).getTagName() + "~*" + childElements.get(i).getText();
						break;
					}
				} 
			}else{
				tagName = "Text"  + "~#" + element.get(j).getTagName() + "~*" + element.get(j).getText();
			}
			break;
		}
		
	}
	return tagName;
	
	
}

public static String findParentTag(WebElement element){
	String tagName = null;
	
		System.out.println(element.getTagName() + " - "+ element.getText());
		if (element.getAttribute("for") != null){
			tagName =  "FOR"  + "~#" + element.getTagName() + "~*" + element.getAttribute("for") ;
			System.out.println(element.getText());
		}/*else if (!element.getText().equals("")){
			System.out.println(element.findElement(By.xpath("descendant::*")).getTagName());
			
			List<WebElement> childElements = element.findElements(By.xpath("descendant::*"));
			if (childElements.size() >0){
				for(int i =0; i < childElements.size();i++){
					if (childElements.get(i).getText() != null & !(childElements.get(i).getText().equals(""))){
						tagName = "Text-decendent"  + "~#" + childElements.get(i).getTagName() + "~*" + childElements.get(i).getText();
						break;
					}
					
				}
			}else{
				tagName = "Text"  + "~#" + element.getTagName() + "~*" + element.getText();
			}*/
			
		
		//}
		
	
	return tagName;
	
	
}

public static String generateXpath(String attribute, String label, String value){
	String xpathValue = null;
	switch (attribute) {
	case "FOR":
		xpathValue = "//" + label +"[@for='" + value +"']";
		break;
	case "Text":
		if (value.length() > 25){
			value = value.substring(0, value.indexOf("\n"));
		}
		xpathValue = "//" + label +"[text()= '" + value +"')]";
		break;
		
	case "Text-decendent":
		if (value.length() > 25){
			value = value.substring(0, value.indexOf("\n"));
			xpathValue = "//" + label +"[contains(text(), '" + value +"')]/..";
		}else{
			xpathValue = "//" + label +"[text()= '" + value +"']/..";
		
		}
		break;
		
	default:
		break;
	}
	
	
	return xpathValue;
	
}

	
}
