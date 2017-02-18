import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.PdfTextExtractor;


public class ResumeParsingTest {
	public static final String RESUME_PARSE_READ_DIRECTORY = "C:/Users/JAYADEEP/Desktop/45_resumes";  
	public static final String RESUME_PARSE_WRITE_DIRECTORY = "C:/Users/JAYADEEP/Desktop/Output/ResumeParsing.xlsx";  
	public static boolean isFound = false; ;
	public static final String FIRST_NAME = "FIRST NAME";
	public static final String LAST_NAME = "LAST NAME";
	public static final String MIDDLE_NAME = "MIDDLE NAME";
	public static final String EMAIL = "EMAIL";
	public static final String PHONE_NUMBER = "PHONE NUMBER";
	public static final String ADDRESS = "ADDRESS";
	public static final String ALTERNATE_NUMBER = "ALTERNATE NUMBER";
	public static final String LINKEDIN_URL = "LINKED-IN URL";
	public static final String OTHER_ONLINE_URL = "OTHER ONLINE URL";
	static int pdfFormat;
	static int docFormat;
	static int docxFormat;
	static int unReadFormat;
	public static ArrayList unReadFormatList = null;
	
	public static void main(String[] args) throws Exception{
		File directory = new File(RESUME_PARSE_READ_DIRECTORY);
		File[] fileList = directory.listFiles();
		unReadFormatList = new ArrayList();
		try {
        if(fileList!=null) {
          //Create blank workbook
  	      XSSFWorkbook workbook = new XSSFWorkbook(); 
  	      //Create a blank sheet
  	      XSSFSheet spreadsheet = workbook.createSheet(" Resume Info ");
  	      //Create row object
	      XSSFRow row;
	      Map < String, Object[] > resuemeInfo =  new TreeMap < String, Object[] >();
	      int rowNum=1 ;
	      resuemeInfo.put( Integer.toString(rowNum++), new Object[] {"File Name",FIRST_NAME, MIDDLE_NAME, LAST_NAME, EMAIL, PHONE_NUMBER, ALTERNATE_NUMBER, ADDRESS, LINKEDIN_URL, OTHER_ONLINE_URL });
        	for (File file : fileList) {
        		String email="";
        		String alternateEmail="";
        		String name="";
        		String firstName="";
        		String lastName="";
        		String middleName="";
        		String address="";
        		String phoneNumber="";
        		String alternatePhoneNumber = "";
        		String linkedInURL = "";
        		String otherOnlineURL = "";
        		
        		String fileName = file.getName();
        		String extension = fileName.substring(fileName.lastIndexOf(".") + 1);
        		
        		//Getting the Page Content
        		String pageContent = getPageContent(file,extension);
        		System.out.println("pageContent >>"+pageContent);
        		
        		ArrayList pageContetInList = new ArrayList();
        		String[] lines = null;
        		if(pageContent != null && !pageContent.isEmpty()) {
        			System.out.println("Valid Format of Resume >>>"+fileName);
        			lines  = pageContent.split("\n");
        		} else {
        			System.out.println("Invalid Format of Resume >>>"+fileName);
        			continue;
        		}
        		
        		pageContetInList.addAll(Arrays.asList(lines));
        		
        		//Filtered the Page Content
        		ArrayList filteredConentList = (ArrayList) filterPageContent(pageContetInList);
        		
        		//Getting the Email Id
        		List<String> emailList = getEmailId(filteredConentList);
        		
        		//Filter Email from the Content
        		if(emailList.size() == 1) {
        			email = emailList.get(0).toString();
        		} else if(emailList.size() >1){
        			email = emailList.get(0).toString();
        			alternateEmail = emailList.get(1).toString();
        		}
        		
        		//Validating Email
        		if(email == null || email.isEmpty()) {
        			System.out.println("Invalid Email >>>"+fileName);
        			//continue;
        		} 
        		
        		name = getName(filteredConentList,email);
        			
        		
        		
        		//Getting the Phone Number
        		//Validating Name
        		if(name == null || name.isEmpty()) {
        			System.out.println("Invalid Name >>>"+fileName);
        		} 
        		
        		List phoneNumList = getPhoneNumber(filteredConentList, name, email);
        		if(phoneNumList.size() == 1) {
        			phoneNumber = phoneNumList.get(0).toString();
        		} else if(phoneNumList.size() == 2) {
        			phoneNumber = phoneNumList.get(0).toString();
        			alternatePhoneNumber = phoneNumList.get(1).toString();
        		}
        		
        		//Getting the Address 
        		List addressList = getAddress(filteredConentList, name, email, phoneNumber, alternatePhoneNumber);
        		address = filterAddress(addressList, name, email, phoneNumber) ;
        		
        		Map url = getURL(address); 
        		if(null != url.get("linkedinUrl"))
        			linkedInURL = url.get("linkedinUrl").toString();
        		if(null != url.get("otherUrl"))
        			otherOnlineURL = url.get("otherUrl").toString();
        		
        		String[] subName = name.split(" ");
        		int subNameLength = subName.length;
        		if(subNameLength == 1) {
        			firstName = subName[0];
        		}else if(subNameLength == 2) {
        			firstName = subName[0];
        			lastName = subName[1];
        		} else if(subNameLength >2) {
        			firstName = subName[0];
        			lastName = subName[subNameLength-1];
        			for(int i=1; i<=subNameLength-2; i++) {
        				middleName += subName[i];
        			}
        		}
        		/*if(!name.isEmpty())
        			System.out.println("Name >>>"+name);
        		if(!firstName.isEmpty())
        			System.out.println("First Name >>>"+firstName);
        		if(!middleName.isEmpty())
        			System.out.println("Middle Name >>>"+middleName);
        		if(!lastName.isEmpty())
        			System.out.println("Last Name >>>"+lastName);
        		if(!email.isEmpty())
        			System.out.println("Email >>>"+email);
        		if(!phoneNumber.isEmpty())
        			System.out.println("Phone Number >>>"+phoneNumber);
        		if(!address.isEmpty())
        			System.out.println("Address >>>"+address);*/
        		
        		//Phone number Refined
        		String phoneNumberRefined = phoneNumber.replaceAll("[^\\d+]", "").trim();
        		String alternateNumberRefined = alternatePhoneNumber.replaceAll("[^\\d+]", "").trim();
        		
        		//Email Refined
        		String emailRefined = email.replaceAll("Email:", "");
        		emailRefined = emailRefined.replaceAll("email:", "");
        		emailRefined = emailRefined.replaceAll("EMAIL:", "");
        		emailRefined = emailRefined.replaceAll(":", "");
        		
        		//Address Refines 
        		/*String addressRefined = "" ;
        		String[] addArr = address.split(",") ;
        		for(String addr : addArr){
        			if(null != addr 
        					&& !addr.contains(firstName) && !addr.contains(middleName) 
        					&& !addr.contains(lastName) && !addr.contains(emailRefined) 
        					&& !addr.contains(phoneNumberRefined) && !addr.contains(alternateNumberRefined)
        					&& !addr.contains(linkedInURL) && !addr.contains(otherOnlineURL)) {
        						addressRefined = addressRefined+" , "+addr;
        					}
        				
        		}*/
        		resuemeInfo.put( Integer.toString(rowNum++), new Object[] {fileName,firstName, middleName, lastName, emailRefined, phoneNumberRefined, alternateNumberRefined, address, linkedInURL, otherOnlineURL });
        	}
        	
        	 //Iterate over data and write to sheet
  	      	Set < String > keyid = resuemeInfo.keySet();
  	      	int rowid = 0;
  	      	for (String key : keyid) {
  	         row = spreadsheet.createRow(rowid++);
  	         Object [] objectArr = resuemeInfo.get(key);
  	         int cellid = 0;
  	         for (Object obj : objectArr)
  	         {
  	            Cell cell = row.createCell(cellid++);
  	            cell.setCellValue((String)obj);
  	         }
  	      }
  	      //Write the workbook in file system
  	      FileOutputStream out = new FileOutputStream(new File(RESUME_PARSE_WRITE_DIRECTORY));
  	      workbook.write(out);
  	      out.close();
  	      System.out.println("Resume Parsing successfully done.." );
  	     /* System.out.println("unReadFormat >>"+unReadFormat);
  	      System.out.println("pdfFormat >>"+pdfFormat);
  	      System.out.println("docFormat >>"+docFormat);
  	      System.out.println("docxFormat >>"+docxFormat);
  	      System.out.println("unReadFormatList"+unReadFormatList.toString());*/
        }
		} catch(Exception e) {
			e.printStackTrace();
		}
		
	}
	
	public static Map getURL(String address) {
		Map map= new LinkedHashMap();
		String[] str = address.split(",");
		for(String line : str) {
			if(null != line) {
				String lineRefined = line.replaceAll("Url", "");
				lineRefined = lineRefined.replaceAll("url", "");
				lineRefined = lineRefined.replaceAll("URL", "");
				lineRefined = lineRefined.replaceAll(":", "");
				if(lineRefined.toLowerCase().contains("linkedin") && line.contains("www")) {
					map.put("linkedinUrl", lineRefined);
				} else if (lineRefined.contains("www.")) {
					map.put("otherUrl", lineRefined);
				}
			}
		}
		return map;
	}
	
	public static String filterAddress(List addressList, String name,String email, String phoneNumber) {
		String address="";
		for(int i=0;i<addressList.size();i++) {
			String line = addressList.get(i).toString().trim();
			if(!line.contains(name) && !line.contains(email) && !line.contains(phoneNumber)) {
				if(address.isEmpty()) {
					address += line;
				} else{
					address = address+" , "+line;
				}
			}
		}
		return address;
	}
	
	public static List getAddress(List pageContentInList, String name,String email, String phoneNumber, String altPhNo) {
		ArrayList al2Address = new ArrayList();
		if(isFound == true)
			pageContentInList.remove(0);
		for(int i=0; i<pageContentInList.size(); i++) {
			String line = pageContentInList.get(i).toString().trim();
			
			if(line.contains(name) && !line.equals(name)) {
				//System.out.println("name checking");
				String[] str= line.split("   ");
				if(str.length == 1) {
					str = line.split("\t\t\t");
				}
				ArrayList al1 = new ArrayList();
				al1.addAll(Arrays.asList(str));
				for(int j=0;j<al1.size();j++) {
					String str1 = al1.get(j).toString();
					if(!str1.isEmpty() && !str1.equals(name) && !str1.contains(email) && !str1.equals(phoneNumber) && !str1.equals(altPhNo)) {
						String addr = str1.replace(",", " ");
						al2Address.add(addr);
					}
				}
			}
			
			if(line.contains(email) && !line.equals(email)) {
				//System.out.println("Email checking");
				String[] str= line.split("   ");
				if(str.length == 1) {
					str = line.split("\t\t\t");
				}
				ArrayList al1 = new ArrayList();
				al1.addAll(Arrays.asList(str));
				for(int j=0;j<al1.size();j++) {
					String str1 = al1.get(j).toString();
					if(!str1.isEmpty() && !str1.equals(name) && !str1.contains(email) && !str1.equals(phoneNumber) && !str1.equals(altPhNo)) {
						String addr = str1.replace(",", " ");
						al2Address.add(addr);
					}
				}
			}
			
			if((line.contains(phoneNumber) && !line.equals(phoneNumber)) || (line.contains(altPhNo) && !line.equals(altPhNo))) {
				//System.out.println("Phone checking");
				String[] str= line.split("   ");
				if(str.length == 1) {
					str = line.split("\t\t\t");
				}
				ArrayList al1 = new ArrayList();
				al1.addAll(Arrays.asList(str));
				for(int j=0;j<al1.size();j++) {
					String str1 = al1.get(j).toString();
					if(!str1.isEmpty() && !str1.equals(name) && !str1.contains(email) && !str1.equals(phoneNumber) && !str1.equals(altPhNo)) {
						String addr = str1.replace(",", " ");
						al2Address.add(addr);
					}
				}
			}
			
			if(!line.isEmpty() && !line.contains(name) && !line.contains(email) && !line.contains(phoneNumber) && !line.contains(altPhNo)) {
				String addr = line.replace(",", " ");
				al2Address.add(addr);
			}
		}
		
		// Remove duplicates variable 
		LinkedHashSet<String> hs = new LinkedHashSet<>();
		hs.addAll(al2Address);
		al2Address.clear();
		al2Address.addAll(hs);
		return al2Address;
	}
	
	public static List getPhoneNumber(List pageContentInList,String name,String email) {
		ArrayList phoneNumber = new ArrayList();
		Pattern p2 = Pattern.compile("\\d"); 
		for(int i=0;i <pageContentInList.size(); i++ ){
			String line = pageContentInList.get(i).toString().trim();
			if(null != email && !email.isEmpty() && line.contains(email) && !line.equals(email)) {
				//System.out.println("Email checking");
				String[] str= line.split("   ");
				if(str.length == 1) {
					str = line.split("\t\t\t");
				}
				ArrayList al1 = new ArrayList();
				al1.addAll(Arrays.asList(str));
				ArrayList al2 = new ArrayList();
				for(int j=0;j<al1.size();j++) {
					String str1 = al1.get(j).toString().trim();
					if(!str1.isEmpty()) {
						al2.add(str1);
					}
				}
				
				for(int j=0;j<al2.size();j++) {
					String str1 = al2.get(j).toString().trim();
					Matcher m = p2.matcher(str1);
					int count = 0;
					while(m.find()){
						count++;
					}
					if(count > 9) {
						phoneNumber.add(str1);
						//break;
					} 
				}
			} else {
				Matcher m2 = p2.matcher(line);
				int count = 0;
				while(m2.find()){
					count++;
				}
				
				if(count > 9) {
					int length = line.length();
					if(length > 16) {
						String[] str= line.split("   ");
						if(str.length == 1) {
							str = line.split("\t\t\t");
						}
						ArrayList al1 = new ArrayList();
						al1.addAll(Arrays.asList(str));
						ArrayList al2 = new ArrayList();
						for(int j=0;j<al1.size();j++) {
							String str1 = al1.get(j).toString().trim();
							if(!str1.isEmpty()) {
								al2.add(str1);
							}
						}
						for(int j=0;j<al2.size();j++) {
							String str1 = al2.get(j).toString().trim();
							Matcher m3 = p2.matcher(str1);
							int count2 = 0;
							while(m3.find()){
								count2++;
							}
							if(count2 > 9) {
								phoneNumber.add(str1);
								//break;
							}
						}
					} else {
						phoneNumber.add(line);
					}
					//break;
				}
			}
			
		}
		if(phoneNumber.size() == 1 && phoneNumber.get(0).toString().trim().length() >40 ) {
			String line = phoneNumber.get(0).toString().trim();
			phoneNumber.clear();
			
			
			String[] str= line.split(" ");
			if(str.length == 1) {
				str = line.split("\t");
			}
			ArrayList al1 = new ArrayList();
			al1.addAll(Arrays.asList(str));
			ArrayList al2 = new ArrayList();
			for(int j=0;j<al1.size();j++) {
				String str1 = al1.get(j).toString().trim();
				if(!str1.isEmpty()) {
					al2.add(str1);
				}
			}
			
			for(int j=0;j<al2.size();j++) {
				String str1 = al2.get(j).toString().trim();
				Matcher m = p2.matcher(str1);
				int count = 0;
				while(m.find()){
					count++;
				}
				if(count > 9) {
					phoneNumber.add(str1);
					break;
				} 
			}

			
		}
		return phoneNumber;
	}
	
	public static List<String> getEmailId(List pageContetInList) { 
		ArrayList emailList = new ArrayList();
		for(int i=0;i<pageContetInList.size();i++) {
			Pattern p =Pattern.compile("\\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,4}\\b",Pattern.CASE_INSENSITIVE);
			Matcher matcher = p.matcher(pageContetInList.get(i).toString());
			while(matcher.find()) {
				String[] str = pageContetInList.get(i).toString().split(" ");
				for(String emailStr : str) {
					Matcher exactMatcher = p.matcher(emailStr);
					while(exactMatcher.find()) {
						emailList.add(emailStr.trim());
					}
				}
			}
		}
		return emailList;
	}
	
	public static String getName(List pageContetInList,String Email) {
		String[] firstLineArr = {"CURRICULAM","VITAE","RESUME","BIO","BIOGRAPHY","PAGE","Type","[","]"};
		isFound = false;
		String name="";	
			if(pageContetInList.size() >0) {
				String resumeFirstLine = pageContetInList.get(0).toString();
				String resumeFirstLineUpper = resumeFirstLine.toUpperCase();
				for(String firstLine : firstLineArr) {
					String firstLineUpper = firstLine.toUpperCase();
					if(resumeFirstLineUpper.contains(firstLineUpper)) {
						isFound  = true;
						break;
					} 
				}
				if(isFound && pageContetInList.size() > 1) {
					name = pageContetInList.get(1).toString().trim();
				} else {
					name = pageContetInList.get(0).toString().trim();
				}
			}
		return name; 
	}
	
	public static List filterPageContent(List pageContetInList) {
		String[] headersArr = {"SUMMARY","SKILLS","EXPERIENCE","PROFILE","CAREER","PERSONAL","SYNOPSIS","PROFESSIONAL","OBJECTIVE","QUALIFICATION","EDUCATION"};
		ArrayList filteredConentList = new ArrayList();
		boolean isProceed = true;
		for(int i=0;i<pageContetInList.size();i++) {
			String resumeLine = pageContetInList.get(i).toString();
			String resumeLineUpper = resumeLine.toUpperCase();
			for(String header : headersArr) {
				String headerUpper = header.toUpperCase();
				if(resumeLineUpper.contains(headerUpper)) {
					isProceed = false;
					break;
				} 
			}
			if(isProceed == false) {
				break;
			} else if(!resumeLine.trim().isEmpty()) {
				filteredConentList.add(resumeLine);
			}
			
		}
		return filteredConentList;
	}
	
	public static String getPageContent(File file,String extension) {
		
		String pageContent=null ;
		if(extension.equalsIgnoreCase("doc")) {
			docFormat++;
			pageContent = readDocFile(file);
		} else if(extension.equalsIgnoreCase("docx")) {
			docxFormat++;
			pageContent =readDocxFile(file);
		} else if(extension.equalsIgnoreCase("pdf")) {
			pdfFormat++;
			pageContent = readPdfFile(file);
		} else {
			unReadFormat++;
			unReadFormatList.add(extension);
		}
		return pageContent;
	}

	public static String readDocFile(File file) {
		String pageContent=null;
		try {
			//System.out.println(">>>Reading Doc File.>>>");
			HWPFDocument doc = new HWPFDocument(new FileInputStream(file));
			WordExtractor we = new WordExtractor(doc);
			pageContent = we.getText();
		} catch(Exception e) {
			System.out.println("Error while reading .doc file");
			e.printStackTrace();
		}
		return pageContent;
	}
	
	public static String readDocxFile(File file) {
		String pageContent=null;
		try {
			//System.out.println(">>>Reading Docx File.>>>");
			XWPFDocument docx = new XWPFDocument(new FileInputStream(file));
			XWPFWordExtractor we = new XWPFWordExtractor(docx);
			pageContent = we.getText();
		} catch(Exception e) {
			System.out.println("Error while reading .docx file");
			e.printStackTrace();
		}
		return pageContent;
	}
	
	public static String readPdfFile(File file) {
		String pageContent="";
		try {
			//System.out.println(">>>Reading PDF File.>>>");
			PdfReader pdfReader = new PdfReader(file.getAbsolutePath());
			int pages = pdfReader.getNumberOfPages(); 
			for(int i=1; i<=pages; i++) { 
				pageContent += PdfTextExtractor.getTextFromPage(pdfReader, i);
			}		
		} catch(Exception e) {
			System.out.println("Error while reading pdf file>>"+file.getName());
			e.printStackTrace();
		}
		return pageContent;
	}
}


