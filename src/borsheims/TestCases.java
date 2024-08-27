package borsheims;

import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import static org.testng.Assert.assertEquals;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.WindowType;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class TestCases {

	WebDriver driver;
	WebDriverWait wait;

	@BeforeTest
	public void setUp() throws Exception {

		ChromeOptions options = new ChromeOptions();
		options.addArguments("--start-maximized");
		options.addArguments("--disable-notifications");

		driver = new ChromeDriver(options);

		wait = new WebDriverWait(driver, Duration.ofSeconds(10));

	}

	@Test
	public void CaseOne() throws Exception {
		// Open the web-site
		driver.get("https://www.borsheims.com/");

		// Click 'Account' link
		WebElement accountLink = driver
				.findElement(By.cssSelector("div.c-cta-style-extra-small a[href*='/customer-account']"));
		accountLink.click();

		wait.until(
				ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"js-main-content\"]/div/div[1]/h1")));

		// Enter valid user information and login
		driver.findElement(By.cssSelector("#l-Customer_LoginEmail")).sendKeys("luna@gmail.com");
		driver.findElement(By.cssSelector("#l-Customer_Password")).sendKeys("luna12345");
		WebElement submitButton = driver.findElement(By.cssSelector("input[type='submit'][value='Login']"));
		submitButton.click();
		// Make sure it's logged in successfully
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("a[title='Log Out']")));
		System.out.println("Logged in successfully");

		// Create a list to store the brands' names
		List<String> brandsList = new ArrayList<>();

		// Read from Excel sheet
		File myFile = new File("C:\\Users\\HITECH\\Desktop\\borsh.xlsx");
		FileInputStream fis = new FileInputStream(myFile);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sh = wb.getSheetAt(1);
		int rowsNum = sh.getLastRowNum();

		// Read the brands' names from the Excel sheet and add them to the list
		for (int i = 0; i <= rowsNum; i++) {
			String brandName = sh.getRow(i).getCell(0).toString();
			brandsList.add(brandName);
		}

		// Select a brand randomly
		Random rnd = new Random();
		int rndIndex = rnd.nextInt(brandsList.size());
		String randomBrand = brandsList.get(rndIndex);

		// Find the search input field and type the brand name chosen randomly
		WebElement searchBar = driver.findElement(By.cssSelector("#l-desktop-search"));
		searchBar.sendKeys(randomBrand); // randomBrand

		// Click search button
		WebElement searchButton = driver.findElement(By.xpath("//*[@id=\"js-search-form\"]/form/div/button"));
		searchButton.click();

		// Wait a little to ensure the page completely loaded
		Thread.sleep(3000);

		// Ensure the page title contains the chosen brand name
		WebElement title = driver.findElement(By.cssSelector("#js-category-header-title"));
		String pageTitle = title.getText();
		Assert.assertTrue(pageTitle.contains(randomBrand), "The page title is wrong"); 
		System.out.println("Brand title: <" + randomBrand + "> is displayed successfully.");
		
		List<WebElement> itemsList = new ArrayList<>();

		int iterator = 3;
		for (int i = 0; i < iterator; i++) {
			try {
				WebElement container = driver
						.findElement(By.xpath("//*[@id=\"js-main-content\"]/div[2]/div[1]/div/div[3]/x-search-app"));
				itemsList = container.getShadowRoot().findElements(By.cssSelector("x-card"));

				Thread.sleep(1000);

				int randomIndex = rnd.nextInt(itemsList.size());
				WebElement product = itemsList.get(randomIndex);

				product.click();

				// Locate the 'addItem' element again if needed
				WebElement addItem = driver.findElement(By.xpath("//*[@id=\"js-mini-basket\"]"));

				// Locate the 'clickAdd' element again to avoid stale element exception
				WebElement clickAdd = addItem.findElement(By.xpath("//*[@id='js-add-to-cart']"));

				wait.until(ExpectedConditions.elementToBeClickable(clickAdd));
				clickAdd.click();

				// Locate the bread-crumbs element again to avoid stale element exception
				WebElement breadcrumbs = wait.until(
						ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"js-prod-breadcrumbs\"]/li[2]/a")));
				breadcrumbs.click();

				Thread.sleep(1000);

				// Ensure the page title contains the chosen brand name
				WebElement title1 = wait.until(
						ExpectedConditions.visibilityOfElementLocated(By.cssSelector("#js-category-header-title")));
				String pageTitle1 = title1.getText();
				Assert.assertTrue(pageTitle1.contains(randomBrand), "The page title is wrong");
				System.out.println("Item " + (i + 1) + " has been added to the bag.");

			} catch (StaleElementReferenceException e) {
				System.out.println("Stale element reference exception caught. Retrying...");
				if (i == iterator) {
					continue;
				}
				i--;
			} catch (Exception e) {
				System.out.println("Exception caught: " + e.getMessage());
			}
		}

		driver.get("https://www.borsheims.com/your-bag");
		Thread.sleep(2000);
		System.out.println("The bag was displayed successfully.");

		fis.close();
		wb.close();

	}

	@Test
	public void CaseTwo() throws Exception {
		// Read from Excel sheet
		File myFile2 = new File("C:\\Users\\HITECH\\Desktop\\borsh.xlsx");
		FileInputStream fis2 = new FileInputStream(myFile2);
		XSSFWorkbook wb2 = new XSSFWorkbook(fis2);
		XSSFSheet sh2 = wb2.getSheetAt(0);
		int rowsNum2 = sh2.getLastRowNum();

		// start rows loop
		for (int i = 0; i < rowsNum2 + 1; i++) {
			XSSFRow row = sh2.getRow(i);
			int colsNum2 = row.getLastCellNum();

			// Register if user is new
			if (sh2.getRow(i).getCell(12) == null) {
				XSSFCell cell = sh2.getRow(i).createCell(12);

				// Open a new tab and switch to it
				driver.switchTo().newWindow(WindowType.TAB);
				assertEquals("", driver.getTitle());
				driver.get("https://www.borsheims.com/");

				// Click 'Account' link
				WebElement accountLink = driver
						.findElement(By.cssSelector("div.c-cta-style-extra-small a[href*='/customer-account']"));
				accountLink.click();

				wait.until(ExpectedConditions.textToBePresentInElementLocated(
						By.xpath("//*[@id=\"js-main-content\"]/div/div[1]/h1"), "Customer Log In"));

				// Click "Register" link
				WebElement registerLink = driver.findElement(By.cssSelector("a[title=\"Register\"]"));
				registerLink.click();

				wait.until(ExpectedConditions.textToBePresentInElementLocated(
						By.xpath("//*[@id=\"js-main-content\"]/div/div[1]/h1"), "Create An Account"));

				// Store the log in email, the log in password, and the user name
				String logInEmail = row.getCell(0).toString();
				String logInPassword = row.getCell(1).toString();
				String username = row.getCell(3).toString();

				int j = 0;
				while (j < colsNum2) {
					// Enter valid user information and register
					driver.findElement(By.cssSelector("#Customer_LoginEmail"))
							.sendKeys(row.getCell(j++).getStringCellValue()); // Sign In Email
					driver.findElement(By.cssSelector("#l-Customer_Password"))
							.sendKeys(row.getCell(j++).getStringCellValue()); // Sign In Password
					driver.findElement(By.cssSelector("#l-Customer_VerifyPassword"))
							.sendKeys(row.getCell(j++).getStringCellValue()); // Verify Password
					driver.findElement(By.cssSelector("#l-Customer_ShipFirstName"))
							.sendKeys(row.getCell(j++).getStringCellValue()); // First Name
					driver.findElement(By.cssSelector("#l-Customer_ShipLastName"))
							.sendKeys(row.getCell(j++).getStringCellValue()); // Last Name
					driver.findElement(By.cssSelector("#Customer_ShipEmail"))
							.sendKeys(row.getCell(j++).getStringCellValue()); // Ship Email
					driver.findElement(By.cssSelector("#l-Customer_ShipPhone"))
							.sendKeys(row.getCell(j++).getStringCellValue()); // Phone Number
					driver.findElement(By.cssSelector("#l-Customer_ShipCompany"))
							.sendKeys(row.getCell(j++).getStringCellValue()); // Company
					driver.findElement(By.cssSelector("#l-Customer_ShipAddress1"))
							.sendKeys(row.getCell(j++).getStringCellValue()); // Ship Address
					driver.findElement(By.cssSelector("#l-Customer_ShipCity"))
							.sendKeys(row.getCell(j++).getStringCellValue()); // Ship City

					// Select The State
					String state = row.getCell(j++).getStringCellValue();
					WebElement stateSelect = driver.findElement(By.cssSelector("#Customer_ShipStateSelect"));
					Select select = new Select(stateSelect); // Create a Select object
					select.selectByVisibleText(state); // Select the state by visible text

					// Convert the numeric zip code to a String
					String zipCode = String.valueOf((long) row.getCell(j++).getNumericCellValue());
					// Send the zip code
					driver.findElement(By.cssSelector("#l-Customer_ShipZip")).sendKeys(zipCode);
				}

				// Check the check box- Bill to: same as shipping
				WebElement checkbox = driver.findElement(By.cssSelector("#billing_to_show"));
				if (!checkbox.isSelected()) {
					checkbox.click();
				}

				// Save registration information
				WebElement submitButton = driver.findElement(By.cssSelector("input[type='submit'][value='Save']"));
				submitButton.click();

				// Wait until the page is loaded successfully
				// Ensure the user name is displayed
				wait.until(ExpectedConditions
						.textToBePresentInElementLocated(By.cssSelector("p.message.message-success"), username));

				System.out.println("Username < " + username + " > is Displayed.");

				WebElement myAccount = driver.findElement(By.cssSelector("a[title=\"Account\"]"));
				myAccount.click();

				wait.until(ExpectedConditions.textToBePresentInElementLocated(
						By.xpath("//*[@id=\"js-main-content\"]/div/div[1]/h1"), "My Account"));

				// Log out
				WebElement logOutButton = driver.findElement(By.cssSelector("a[title='Log Out']"));
				logOutButton.click();

				// Wait until redirected to the main page
				Thread.sleep(5500);

				// Log in again
				WebElement accountLink2 = driver
						.findElement(By.cssSelector("div.c-cta-style-extra-small a[href*='/customer-account']"));
				accountLink2.click();

				wait.until(ExpectedConditions.textToBePresentInElementLocated(
						By.xpath("//*[@id=\"js-main-content\"]/div/div[1]/h1"), "Customer Log In"));

				driver.findElement(By.cssSelector("#l-Customer_LoginEmail")).sendKeys(logInEmail);
				driver.findElement(By.cssSelector("#l-Customer_Password")).sendKeys(logInPassword);
				WebElement logInButton = driver.findElement(By.cssSelector("input[type='submit'][value='Login']"));
				logInButton.click();

				wait.until(ExpectedConditions.textToBePresentInElementLocated(
						By.xpath("//*[@id=\"js-main-content\"]/div/div[1]/h1"), "My Account"));

				// Log out
				WebElement logOutButton2 = driver.findElement(By.cssSelector("a[title='Log Out']"));
				logOutButton2.click();

				// Check user is registered and type 'X' in the row
				System.out.println("User " + i + " Successfully Registered.");
				cell.setCellValue("X");

				// The cell #12 is not null and the user is already registered
			} else {
				System.out.println("User " + i + " Already Registered");
			}

		} // end rows Loop

		// Write the changes to the workbook
		FileOutputStream wf2 = new FileOutputStream(myFile2);
		wb2.write(wf2);
		wf2.close();
		fis2.close();
		wb2.close();

	}

	@AfterTest
	public void teardown() {
		// Quit after finishing the test
		if (driver != null) {
			driver.quit();
		}
	}

	public static void main(String[] args) throws Exception {
		TestCases test = new TestCases();
		test.setUp();

		// Read the Excel sheet
		String FilePath = "C:\\Users\\HITECH\\Desktop\\borsh.xlsx";
		File myFile = new File(FilePath);
		FileInputStream fis = new FileInputStream(myFile);
		XSSFWorkbook wbx = new XSSFWorkbook(fis);
		XSSFSheet sh1 = wbx.getSheetAt(2);

		FileOutputStream wf = null;

		// No test case is marked
		if (!sh1.getRow(1).getCell(1).toString().equals("X") && !sh1.getRow(2).getCell(1).toString().equals("X")) {
			System.out.println("No Test Case Marked");
		}

		// Test Case 1
		if (sh1.getRow(1).getCell(0).toString().equals("TC1") && sh1.getRow(1).getCell(1).toString().equals("X")) {
			try {
				test.CaseOne();
				if (sh1.getRow(1).getCell(2) == null) {
					sh1.getRow(1).createCell(2).setCellValue("Passed");
				} else {
					sh1.getRow(1).getCell(2).setCellValue("Passed");
				}
				System.out.println("Test Case 1 Passed. Data is written successfully for Test Case 1");
			} catch (Exception e) {
				if (sh1.getRow(1).getCell(2) == null) {
					sh1.getRow(1).createCell(2).setCellValue("Failed");
				} else {
					sh1.getRow(1).getCell(2).setCellValue("Failed");
				}
				System.out.println("Test Case 1 failed. Data is written successfully for Test Case 1.");
			}
		}

		// Test Case 2
		if (sh1.getRow(2).getCell(0).toString().equals("TC2") && sh1.getRow(2).getCell(1).toString().equals("X")) {
			try {
				test.CaseTwo();
				if (sh1.getRow(2).getCell(2) == null) {
					sh1.getRow(2).createCell(2).setCellValue("Passed");
				} else {
					sh1.getRow(2).getCell(2).setCellValue("Passed");
				}
				System.out.println("Test Case 2 Passed. Data is written successfully for Test Case 2");
			} catch (Exception e) {
				if (sh1.getRow(2).getCell(2) == null) {
					sh1.getRow(2).createCell(2).setCellValue("Failed");
				} else {
					sh1.getRow(2).getCell(2).setCellValue("Failed");
				}
				System.out.println("Test Case 2 failed. Data is written successfully for Test Case 2.");
			}
		}

		// Write to the Excel file
		wf = new FileOutputStream(FilePath);
		wbx.write(wf);

		if (wf != null) {
			wf.close();
		}
		wbx.close();

		test.teardown();

	}
}
