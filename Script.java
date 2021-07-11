package excelExportAndFileIO;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.logging.FileHandler;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;
import org.apache.commons.io.FileUtils;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;

public class Nassau_Re_06_11 {
	
	
	public static void main(String[] args) throws InterruptedException, EncryptedDocumentException, IOException {
		
		int count = 0;
		int failedCount=0;
		int failedLoop=0;
		List<String> FailedCases = Lists.newArrayList();
		List<String> PoliciesFailed = Lists.newArrayList();
		
		//-----------------------------------LogFile Setup-------------------------------------------------------
		
	    Logger logger = Logger.getLogger("MyLog");  
	    FileHandler fh;  

	    try {  

	        // This block configure the logger with handler and formatter  
	        fh = new FileHandler("C:\\Selenium\\LogFile.log");  
	        logger.addHandler(fh);
	        SimpleFormatter formatter = new SimpleFormatter();  
	        fh.setFormatter(formatter);  

	        // the following statement is used to log any messages  
	        logger.info("Execution Started");  

	    } catch (SecurityException e) {  
	        e.printStackTrace();  
	    } catch (IOException e) {  
	        e.printStackTrace();  
	    }  

	
		//-------------------------------------------------------------------------------------------
		//-------------------------------Download PDF Functionality------------------------------------
		String downloadFilepath = "C:\\Selenium\\Proposals";

		//-----------------Creating Random Folder Name------------------------------------------------
		Date date = new Date();
		SimpleDateFormat formatter = new SimpleDateFormat("_dd_MM_HH_mm");
		String strDate= formatter.format(date);
		downloadFilepath=downloadFilepath+strDate;
		new File(downloadFilepath).mkdirs();
       //----------------------------------------------------------------------------------------------
		System.setProperty("webdriver.chrome.driver", "C:\\Selenium\\chromedriver.exe");
		HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("profile.default_content_settings.popups", 0);
		chromePrefs.put("download.default_directory", downloadFilepath);
		chromePrefs.put("plugins.always_open_pdf_externally", true);
		chromePrefs.put("plugins.download.prompt_for_download", false);

		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("prefs", chromePrefs);
		options.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
		options.setCapability(ChromeOptions.CAPABILITY, options);
		WebDriver driver = new ChromeDriver(options);
		driver.manage().window().maximize();		
		
		logger.info("Initial Chrome Setup Completed");
 
		//----------------------------------------------------------------------------------------
	  //-------------------------------Reading Values from Excel sheet----------------------
			Workbook wb;
			Sheet sh;
			FileInputStream fis;
			fis = new FileInputStream("C:\\Selenium\\TestSheet.xlsx");
			wb =WorkbookFactory.create(fis);
			sh = wb.getSheet("Sheet1");
			int noOfRows;
			noOfRows= sh.getPhysicalNumberOfRows();
			List<String> policyData = Lists.newArrayList();
		    for (int i=1; i<noOfRows; i++) {
		        //Print Excel data in console
		    	policyData.add(sh.getRow(i).getCell(0).toString());
		    } 
		    logger.info("Executed function for reading excel file");
	    
	//--------------------------------------------------------------------------------------------
		    int TotalProposal=noOfRows-1;
		    logger.info("Total Number of Proposals present in the excel sheet "+TotalProposal);

	//--------------------------------------------------------------------------------------------
	// Launch website  
		    //driver.navigate().to("http://vm-phx1.navisys.com/phx95u01/integration/test/PHLWMEntry.jsp");
		    driver.navigate().to("http://vm-phx1.navisys.com/phx95cd2/integration/test/PHLWMEntry.jsp");		    
	//--------------------------------------------------------------------------------------------
		    Thread.sleep(3000);     
		    
      //--------------------------------------LOGIN-------------------------------------------------   
		    driver.findElement(By.name("USER")).sendKeys("PHXHO1");
		    driver.findElement(By.name("PW")).sendKeys("1234567");  
		    driver.findElement(By.name("Submit")).click();
		    driver.findElement(By.linkText("Click here to start Illustrations.")).click();

	  //--------------------------------------------------------------------------------------------
		    for (String handle : driver.getWindowHandles()) {
		        driver.switchTo().window(handle);
		    }
		    Thread.sleep(4000);
		    
	  //----------------------------------Close Unnecessary Tab-------------------------------------
		    String originalHandle = driver.getWindowHandle();
		    for(String handle : driver.getWindowHandles()) {
		        if (!handle.equals(originalHandle)) {
		            driver.switchTo().window(handle);
		            driver.close();
		        }
		    }
		    
		    driver.switchTo().window(originalHandle);
	  //--------------------------------------------------------------------------------------------
		        
		        
		    List<String> proposalName = Lists.newArrayList();
	
		    logger.info("Starting Main Loop");
	  //---------------------------Main Loop----------------------------------------------------------------        
		       for(int i =0; i<TotalProposal; i++){
		      
		    driver.findElement(By.id("TABS_PROPOSAL_REPROPOSAL")).click();
		    driver.findElement(By.name("producerCode")).click();
		    driver.findElement(By.name("producerCode")).clear();
		    driver.findElement(By.name("producerCode")).sendKeys("000025");
		    driver.findElement(By.name("proposalName")).click();
		    
		      //-----------------Creating Random Proposal Name------------------------------------------------
			  DateTimeFormatter dtf = DateTimeFormatter.ofPattern("_dd/MM_HH:mm:ss");
			  LocalDateTime now = LocalDateTime.now();
			  String time = dtf.format(now);
			  proposalName.add(policyData.get(i)+time);
			  
			  //--------------------------------------------------------------------------------------			  
			    driver.findElement(By.name("proposalName")).sendKeys(proposalName.get(i));
			    driver.findElement(By.name("accNumber")).click();
			    driver.findElement(By.name("accNumber")).sendKeys(policyData.get(i)); 
			    driver.findElement(By.xpath("//input[2]")).click(); 
			    
				//-------------Open Policy-------------------------------------------------
			    Thread.sleep(3000);
			    driver.findElement(By.xpath("//b[contains(.,'"+proposalName.get(i)+"')]")).click();
			    Thread.sleep(2000);
			    driver.findElement(By.name("create")).click(); 
			    Thread.sleep(2000);
			    
			    //--------------------Close Popup window after download----------------------
			    for(String handle : driver.getWindowHandles()) {
			        if (!handle.equals(originalHandle)) {
			            driver.switchTo().window(handle);
			            driver.close();
			        }
			    }
			    
			    driver.switchTo().window(originalHandle);
			    
			    //-----------------------------------------------------------------------------
			    
			    //------------------Rename and check PDF---------------------------------------------
			    Thread.sleep(2000);
		        File oldName = new File(downloadFilepath+"\\pdf.pdf");
		        File newName = new File(downloadFilepath+"\\"+policyData.get(i)+".pdf");
		        String ProposalName = policyData.get(i);
		        
		        if(oldName.renameTo(newName)) {
		          // logger.info("renamed");
		           count++;
		           logger.info("Download Proposal "+count+ " of "+TotalProposal+" ***Successfully Created*** "+newName);
		           logger.info("Count = "+count);
		           
		        } else {
		           if (failedLoop==0) {
		           logger.info("Failed to Create "+newName+". It will be reattempted again.");
   	           	   FailedCases.add(ProposalName);
					//				FailedCases.add(ProposalName);
					//logger.info("Failed to generate "+newName+". It will be reattempted again");
					//System.out.println("Printing the failed values");
					//System.out.println(FailedCases);
					//System.out.println("Value of i "+i);
				   failedCount++;
				   logger.info("Number of proposals failed to generate "+failedCount);
		           }
		           
		           if (failedLoop==1) {
		        	   logger.info("************************************************ERROR*** Failed to Create "+newName);
		        	   failedCount++;
		        	   logger.info("Number of proposals failed to generate"+failedCount);
		        	   PoliciesFailed.add(ProposalName);
		        	   
		        	   }
		         //---------------------------------------------------------------------------------------------	   
		        }

			    //-----------------------------------------------------------------------------	    
			    driver.findElement(By.xpath("//input[@name='cancel']")).click();
			    Thread.sleep(2000);
			    //------------------------------------------------------------------------------
			    
			    
			    //-----------------------------------------Delete Saved Proposal from Website----------------------------------------------------------------------------
			    driver.findElement(By.xpath("//*[@id=\"TABS_ILLUSTRATOR\"]")).click();
			    Thread.sleep(2500);
			    driver.findElement(By.xpath("//*[@id=\"searchResults\"]/table[2]/tbody/tr[2]/td[1]/input[1]")).click();
			    Thread.sleep(2500);
			    driver.findElement(By.xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr/td/form[2]/div/table[1]/tbody/tr[2]/td[1]/input")).click();
			    Thread.sleep(2500);
			    driver.findElement(By.xpath("/html/body/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr/td/ul/form/table[2]/tbody/tr/td/input")).click();
			    Thread.sleep(1000);
			   logger.info("Proposal Deleted from the website");
			    //---------------------------------------------------------------------------------------------------------------------------------------------------------
/*			 System.out.println("Total Proposal "+TotalProposal);
			 System.out.println("i "+i);
			 System.out.println("failedCount "+failedCount);
			 System.out.println("failedLoop "+failedLoop);
*/
			 
		   if (i==TotalProposal-1 && failedCount>=1 && failedLoop==0) {
				   	logger.info("**********Reattempting the Failed Proposals**********");
				  logger.info("Total failed porposals "+failedCount);
				  logger.info("List of the failed proposals "+FailedCases);
				  //Reinitialization
				  i=-1;
				  failedCount=0;
				  TotalProposal=failedCount;
				  policyData=FailedCases;
				  failedLoop++;
			   }
	   
			}

		       for(String handle : driver.getWindowHandles()) {
		           if (!handle.equals(originalHandle)) {
		               driver.switchTo().window(handle);
		               driver.close();
		           }
		       }
		       
		       driver.close();
		    

	  //--------------------------------------------------------------------------------------------
		       logger.info("***************************************Completion Status**********************************");
		       
		       logger.info("Total Proposals Created = "+count);

		       logger.info("Total Policies present = "+TotalProposal);
		       if (count==noOfRows-1) {
		    	   
		    	   logger.info(count+ " Out of " +TotalProposal+" Proposals downloaded Successfully");
		       }
		       else {
		    	   logger.info(count+ " Out of " +TotalProposal+ " have been generated Successfully");
		    	   logger.info("List of proposals failed to download "+PoliciesFailed+ " You may need to downlaod them manually");
		       }
		 	
		       
		       
//-----------------------------Copying Files for Comparison--------------------------------------------------------------
		       
				File source = new File(downloadFilepath);
				File dest = new File("C:\\Selenium\\Beyond Compare Script\\Current_Files");
				try {
					FileUtils.cleanDirectory(dest);
					FileUtils.copyDirectory(source, dest);
					System.out.println("Files Moved Successfully");
					
				} catch (IOException e) {
				    e.printStackTrace();
				}

		       logger.info("All "+count+" files moved to " +dest);
	       
//-----------------------------------------------------------------------------------------------------------------------	       
	    
	    
}
	
	
}
