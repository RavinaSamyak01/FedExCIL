package fedexCILStaging;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class FedExCILOrderCreation {

	static WebDriver driver;
	static StringBuilder msg = new StringBuilder();
	static String jobid;

	@BeforeMethod
	public void login() throws InterruptedException {
		System.setProperty("webdriver.chrome.driver", ".\\chromedriver.exe");

		ChromeOptions options = new ChromeOptions();
		driver = new ChromeDriver(options);
		String baseUrl = "http://10.20.104.82:9077/TestApplicationUtility/CILOrderCreationClient";
		driver.get(baseUrl);

		driver.manage().window().maximize();
		Thread.sleep(5000);

	}

	@Test
	public static void fedEXCILOrder() throws Exception {
		// Read data from Excel
		File src = new File("C:\\Ravina\\FedExCIL\\TestFiles\\FedExCILTestResult.xlsx");
		FileInputStream fis = new FileInputStream(src);
		Workbook workbook = WorkbookFactory.create(fis);
		Sheet sh1 = workbook.getSheet("Sheet1");

		for (int i = 1; i < 11; i++) {
			DataFormatter formatter = new DataFormatter();
			String file = formatter.formatCellValue(sh1.getRow(i).getCell(0));
			// String TFolder=".//TestFiles//";
			driver.findElement(By.id("MainContent_ctrlfileupload"))
					.sendKeys("C:\\Ravina\\FedExCIL\\TestFiles\\" + file + ".txt");
			Thread.sleep(1000);
			driver.findElement(By.id("MainContent_btnProcess")).click();
			Thread.sleep(3000);
			String Job = driver.findElement(By.id("MainContent_lblresult")).getText();

			// System.out.println(Job);

			Pattern pattern = Pattern.compile("\\w+([0-9]+)\\w+([0-9]+)");
			Matcher matcher = pattern.matcher(Job);
			matcher.find();
			jobid = matcher.group();
			System.out.println("JOB# " + jobid);

			File src1 = new File("C:\\Ravina\\FedExCIL\\TestFiles\\FedExCILTestResult.xlsx");
			FileOutputStream fis1 = new FileOutputStream(src1);
			Sheet sh2 = workbook.getSheet("Sheet1");
			sh2.getRow(i).createCell(1).setCellValue(jobid);
			workbook.write(fis1);
			fis1.close();
			msg.append("JOB # " + jobid + "\n");
		}

	}

	@AfterSuite
	public void SendEmail() throws Exception {
		String subject = "STAGING : FedEx_CIL EDI - Shipment Creation Using SELENIUM";
		try {
			//
			Email.sendMail(
					"ravina.prajapati@samyak.com, asharma@samyak.com,parth.doshi@samyak.com,kunjan.modi@samyak.com, pgandhi@samyak.com",
					subject, msg.toString(), "");
		} catch (Exception ex) {
			Logger.getLogger(FedExCILOrderCreation.class.getName()).log(Level.SEVERE, null, ex);
		}
	}

	@AfterTest
	public void Complete() throws Exception {
		driver.close();
	}
}
