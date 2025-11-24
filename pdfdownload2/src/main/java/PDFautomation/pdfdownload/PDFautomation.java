package PDFautomation.pdfdownload;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.Set;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.UnexpectedTagNameException;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;

public class PDFautomation {

    private static final String BASE_URL = "https://ippne.truchart.com/truchart_app/Home.jsp";
    private static final String USERNAME = "CDWadmin1";
    private static final String PASSWORD = "CDWadmin1";
    private static final String DEFAULT_EXCEL = "downloads/CDW - Care Plan Work.xlsx";

    public static void main(String[] args) throws Exception {
        String excelPath = args.length > 0 ? args[0] : DEFAULT_EXCEL;
        Path downloadsDir = Paths.get(System.getProperty("user.home"), "Downloads");

        WebDriverManager.chromedriver().setup();

        ChromeOptions options = new ChromeOptions();
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("plugins.always_open_pdf_externally", true);
        prefs.put("download.prompt_for_download", false);
        prefs.put("download.default_directory", downloadsDir.toString());
        options.setExperimentalOption("prefs", prefs);

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
        driver.manage().window().maximize();

        // ====== ADDED: declare accountNumber here so it is in scope for the whole method ======
        String accountNumber = null;

        try {
            System.out.println("Started (Excel: " + excelPath + ")");
            driver.get(BASE_URL);

            wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("tru_username"))).sendKeys(USERNAME);
            driver.findElement(By.id("j_password")).sendKeys(PASSWORD);
            wait.until(ExpectedConditions.elementToBeClickable(
                    By.xpath("//form[@name='login']//button[.//span[normalize-space()='Login']]")
            )).click();

            // Save the root window handle to return to it after each iteration
            final String rootWindow = driver.getWindowHandle();

            // -------------------------
            // Open Advanced Participant Search modal (magnifier -> advanced)
            // -------------------------

            // ==== TEST SCREENSHOT AFTER LOGIN ====
try {
    // Make sure the folder exists (will create if missing)
    Files.createDirectories(Path.of("C:/screenshots"));

    String filename = "login_test_" + System.currentTimeMillis() + ".png";

    File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
    Files.copy(src.toPath(), Path.of("C:/screenshots", filename));

    System.out.println("Screenshot saved at: C:/screenshots/" + filename);
} catch (Exception e) {
    e.printStackTrace();
    System.out.println("Screenshot capture FAILED.");
}
// ==== END TEST SCREENSHOT ====

            // ====== ADDED: Read up to 50 account numbers from Excel column B (starting from B2) ======
            List<String> accountNumbers = readAccountNumbers(excelPath, /*startRowIndex=*/1, /*colIndex B=*/1, /*maxCount=*/50);
            if (accountNumbers.isEmpty()) {
                throw new RuntimeException("No account numbers found in Excel (column B starting from row 2).");
            }
            System.out.println("Total accounts to process: " + accountNumbers.size());

            // ====== ADDED: iterate each account and run the same flow via a dedicated method ======
            int idx = 0;
            for (String acct : accountNumbers) {
                idx++;
                accountNumber = acct;
                System.out.println("========== Processing " + idx + "/" + accountNumbers.size() + " | Account: " + accountNumber + " ==========");
                processAccount(driver, wait, downloadsDir, accountNumber, rootWindow);
            }

            System.out.println("All accounts processed.");

        } finally {
            driver.quit();
        }
    }

    // ====== ADDED: Single-account flow extracted into a dedicated method (comments retained inside) ======
    private static void processAccount(WebDriver driver, WebDriverWait wait, Path downloadsDir, String accountNumber, String rootWindow) throws Exception {
        // Ensure we are on the root/main window and main DOM before each iteration
        try {
            driver.switchTo().window(rootWindow);
            driver.switchTo().defaultContent();
        } catch (Exception ignore) {}

        // Freshly open the top search area each time (keeps behavior stable across iterations)
        wait.until(ExpectedConditions.elementToBeClickable(By.id("single-search"))).click();
        WebElement advSearchBtn = wait.until(
                ExpectedConditions.presenceOfElementLocated(
                        By.xpath("//button[@title='Advanced Participant Search']")
                )
        );
        // Use JS click for reliability in VDI environments
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", advSearchBtn);

        // Wait for the modal container to appear
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.ui-dialog[role='dialog']")));
        System.out.println("Advanced Search modal opened successfully");

        // -------------------------
        // Switch to iframe, select Membership -> All, enter account number, click Find
        // -------------------------
        boolean switched = false;
        String[] iframeSelectors = new String[] {
                "iframe[id*='Advanced']",
                "iframe[id*='Search']",
                "div.ui-dialog iframe",
                "iframe.ui-dialog-content",
                "iframe"
        };

        for (String sel : iframeSelectors) {
            try {
                if ("iframe".equals(sel)) {
                    List<WebElement> frames = driver.findElements(By.cssSelector("div.ui-dialog[role='dialog'] iframe"));
                    if (frames != null && !frames.isEmpty()) {
                        driver.switchTo().frame(frames.get(0));
                        switched = true;
                        System.out.println("Switched to iframe found inside dialog.");
                        break;
                    }
                    List<WebElement> allFrames = driver.findElements(By.tagName("iframe"));
                    if (!allFrames.isEmpty()) {
                        driver.switchTo().frame(allFrames.get(0));
                        switched = true;
                        System.out.println("Switched to the first iframe found in the page.");
                        break;
                    }
                } else {
                    try {
                        wait.withTimeout(Duration.ofSeconds(5))
                            .until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.cssSelector(sel)));
                        switched = true;
                        System.out.println("Switched to iframe using selector: " + sel);
                        break;
                    } finally {
                        wait.withTimeout(Duration.ofSeconds(30));
                    }
                }
            } catch (Exception e) {
                // ignore and try next selector
            }
        }

        if (!switched) {
            System.out.println("Info: No iframe switch performed; proceeding in main DOM context.");
        }

        try {
            // Select Membership -> All
            By[] membershipCandidates = new By[] {
                    By.xpath("//form[@name='form_patientSearch']//select[@name='membershipstate']"),
                    By.name("membershipstate"),
                    By.xpath("//label[contains(normalize-space(.),'Membership')]/following::select[1]"),
                    By.cssSelector("div.ui-dialog select[name='membershipstate']"),
                    By.xpath("(//select)[position() < 10 and contains(@name,'membership')]")
            };

            WebElement selElem = null;
            for (By candidate : membershipCandidates) {
                try {
                    selElem = wait.until(ExpectedConditions.presenceOfElementLocated(candidate));
                    if (selElem != null) {
                        System.out.println("Found membership select using: " + candidate.toString());
                        break;
                    }
                } catch (Exception ignored) {
                }
            }

            if (selElem == null) {
                throw new RuntimeException("Membership select not found with any candidate locator.");
            }

            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'}); arguments[0].focus();", selElem);

            boolean selected = false;
            try {
                Select dropdown = new Select(selElem);
                dropdown.selectByVisibleText("All");
                selected = true;
                System.out.println("Selected 'All' using native Select.");
            } catch (Exception e) {
                try {
                    selElem.click();
                    WebElement opt = selElem.findElement(By.xpath(".//option[normalize-space()='All']"));
                    new Actions(driver).moveToElement(opt).click().perform();
                    selected = true;
                    System.out.println("Selected 'All' using Actions click on option.");
                } catch (Exception ignore) {
                    // fall through
                }
            }

            if (!selected) {
                ((JavascriptExecutor) driver).executeScript(
                        "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
                        selElem, "All");
                System.out.println("Selected 'All' using JS fallback.");
            }

            // small pause for handlers
            try {
                Thread.sleep(200);
            } catch (InterruptedException ie) {
                Thread.currentThread().interrupt();
            }

            // ====== CHANGED: assign to the already-declared accountNumber variable ======
            // accountNumber read in caller loop; do not overwrite here

            // Find and fill the "Acct #" field (still in iframe context)
            By[] accountFieldCandidates = new By[] {
                    By.xpath("//form[@name='form_patientSearch']//input[@name='accountnumber']"),
                    By.name("accountnumber"),
                    By.xpath("//label[contains(normalize-space(.),'Acct')]/following::input[1]"),
                    By.cssSelector("input[name='accountnumber']"),
                    By.xpath("//input[@type='text' and contains(@name,'account')]")
            };

            WebElement accountField = null;
            for (By candidate : accountFieldCandidates) {
                try {
                    accountField = wait.until(ExpectedConditions.presenceOfElementLocated(candidate));
                    if (accountField != null) {
                        System.out.println("Found account field using: " + candidate.toString());
                        break;
                    }
                } catch (Exception ignored) {
                }
            }

            if (accountField == null) {
                throw new RuntimeException("Account number field not found with any candidate locator.");
            }

            accountField.clear();
            accountField.sendKeys(accountNumber);
            System.out.println("Entered account number: " + accountNumber);

            // Click Find (scoped to iframe)
            By findButton = By.xpath("//button[normalize-space()='Find' or normalize-space()='Search'] | //button[@id='FindButton']");
            WebElement findBtn = wait.until(ExpectedConditions.elementToBeClickable(findButton));
            findBtn.click();
            System.out.println("Clicked Find.");

            // ===============================
            // After Find: still inside iframe!
            // ===============================

            WebDriverWait shortWait = new WebDriverWait(driver, Duration.ofSeconds(10));

            By resultsRowsSelector = By.cssSelector(
                    "table#patientSearchTablesorter tbody tr.patient-result-row, " +
                    "table.tablesorter tbody tr.patient-result-row, " +
                    "table.tablesorter tbody tr"
            );

            // Wait for table rows inside iframe
            List<WebElement> rows = shortWait.until(
                    ExpectedConditions.numberOfElementsToBeMoreThan(resultsRowsSelector, 0)
            );

            System.out.println("Result rows found: " + rows.size());

            // Click the first result row
            WebElement firstRow = rows.get(0);
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", firstRow);
            firstRow.click();

            System.out.println("Clicked first search result row.");

            // NOW switch back to main window
            driver.switchTo().defaultContent();
            System.out.println("Switched back to default content after clicking result row.");

        } finally {
            // Switch back to default content
            try {
                driver.switchTo().defaultContent();
                System.out.println("Switched back to default content.");
            } catch (Exception e) {
                System.out.println("Warning: Could not switch back to defaultContent(): " + e.getMessage());
            }
        }

        // After Find, wait for results and click first row (back in main DOM)
       // wait.until(ExpectedConditions.elementToBeClickable(
        //        By.cssSelector("table.tablesorter tbody tr"))).click();

        // Continue rest of flow unchanged
        wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//span[@class='quick-tool-desc' and normalize-space()='CarePlan']"))).click();
        switchToNewWindow(driver);

        // ----------------------
        // Robust: find modal -> iframe -> click IDT
        // ----------------------
        System.out.println("DEBUG: waiting for the modal dialog to appear...");
        WebElement dialog = null;
        try {
            dialog = wait.until(ExpectedConditions.visibilityOfElementLocated(
                    By.cssSelector("div.ui-dialog[role='dialog'], div.ui-dialog.ui-widget")));
            System.out.println("DEBUG: dialog found (title: " + dialog.getAttribute("innerText").replaceAll("\\s+", " ").trim().substring(0, Math.min(80, dialog.getAttribute("innerText").length())) + ")");
        } catch (Exception e) {
            System.out.println("DEBUG: modal dialog not found: " + e.getMessage());
            throw new RuntimeException("Modal dialog missing - cannot proceed to quickview");
        }

        // Wait for overlay disappear (if any)
        try {
            WebDriverWait shortWait2 = new WebDriverWait(driver, Duration.ofSeconds(10));
            shortWait2.until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector(".ui-widget-overlay, .loading, .modal-backdrop")));
        } catch (Exception ignore) { /* continue even if overlay persists briefly */ }

        // Try to find iframe inside the dialog (most reliable)
        WebElement iframeInDialog = null;
        try {
            iframeInDialog = dialog.findElement(By.tagName("iframe"));
            System.out.println("DEBUG: iframe found inside dialog via dialog.findElement(tagName('iframe'))");
        } catch (Exception e) {
            System.out.println("DEBUG: no iframe directly under dialog: " + e.getMessage());
        }

        // If not found, fallback to any iframe in the page (we already know there's 1)
        if (iframeInDialog == null) {
            List<WebElement> allFrames = driver.findElements(By.tagName("iframe"));
            System.out.println("DEBUG: page iframe count = " + allFrames.size());
            if (!allFrames.isEmpty()) {
                // choose the first iframe that contains quickview content by inspecting its innerText via JS
                for (int i = 0; i < allFrames.size(); i++) {
                    WebElement f = allFrames.get(i);
                    try {
                        String txt = (String)((JavascriptExecutor) driver).executeScript(
                                "try { return arguments[0].contentDocument ? arguments[0].contentDocument.body.innerText : ''; } catch(e) { return ''; }",
                                f);
                        if (txt != null && txt.toLowerCase().contains("lifeplan") || (txt != null && txt.toLowerCase().contains("idt"))) {
                            iframeInDialog = f;
                            System.out.println("DEBUG: selected iframe index " + i + " by inspecting innerText");
                            break;
                        }
                    } catch (Exception ex) {
                        // ignore cross-origin issues, just continue
                    }
                }
                // if still null, fallback to first iframe
                if (iframeInDialog == null) {
                    iframeInDialog = allFrames.get(0);
                    System.out.println("DEBUG: fallback to first iframe (index 0)");
                }
            } else {
                throw new RuntimeException("No iframes found on the page and none inside the dialog.");
            }
        }

        // Now switch into the iframe by WebElement (reliable)
        try {
            driver.switchTo().frame(iframeInDialog);
            System.out.println("DEBUG: switched into the selected iframe");
        } catch (Exception e) {
            throw new RuntimeException("Failed to switch into iframe: " + e.getMessage(), e);
        }

        // Give the frame a short moment to settle and ensure no overlays inside frame
        try { Thread.sleep(200); } catch (InterruptedException ie) { Thread.currentThread().interrupt(); }
        try {
            new WebDriverWait(driver, Duration.ofSeconds(10))
                .until(ExpectedConditions.invisibilityOfElementLocated(By.cssSelector(".ui-widget-overlay, .loading")));
        } catch (Exception ignore) {}

        // Re-locate the IDT element inside the frame (always re-find to avoid staleness)
        WebElement idtElement = null;
        By[] idtCandidates = new By[] {
            By.id("lp_quickview_3_center"),
            By.xpath("//li[contains(@id,'lp_quickview') and .//text()[contains(translate(., 'idt','IDT'),'IDT')]]"),
            By.xpath("//li[.//span and contains(translate(normalize-space(.//span/text()), 'idt','IDT'),'IDT')]"),
            By.xpath("//span[translate(normalize-space(.), 'idt','IDT')='IDT' or translate(normalize-space(.), 'idt','IDT')='Idt' or normalize-space(.)='Idt']")
        };

        for (By cand : idtCandidates) {
            try {
                idtElement = new WebDriverWait(driver, Duration.ofSeconds(6))
                        .until(ExpectedConditions.presenceOfElementLocated(cand));
                if (idtElement != null) {
                    System.out.println("DEBUG: found IDT candidate using: " + cand);
                    break;
                }
            } catch (Exception ignored) {}
        }

        if (idtElement == null) {
            // last diagnostic: dump some snippet of frame body so you can paste for further analysis
            String frameText = "";
            try {
                frameText = (String) ((JavascriptExecutor) driver).executeScript("return document.body ? document.body.innerText.substring(0,500) : '';");
            } catch (Exception ex) { frameText = "ERROR reading frame text: " + ex.getMessage(); }
            System.out.println("DEBUG: couldn't find IDT inside frame. frameText snippet: " + frameText);
            driver.switchTo().defaultContent();
            throw new RuntimeException("IDT element not found inside frame. See debug output above.");
        }

        // Scroll into view and attempt click (use JS fallback)
        try {
            ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView({block:'center'});", idtElement);
            // small pause
            try { Thread.sleep(150); } catch (InterruptedException ie) { Thread.currentThread().interrupt(); }
            new WebDriverWait(driver, Duration.ofSeconds(6)).until(ExpectedConditions.elementToBeClickable(idtElement));
            idtElement.click();
            System.out.println("DEBUG: clicked IDT using WebElement.click()");
        } catch (Exception e) {
            System.out.println("DEBUG: WebElement.click() failed: " + e.getMessage() + " — trying JS click.");
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", idtElement);
            System.out.println("DEBUG: clicked IDT using JS fallback");
        }

        // return to default content for the rest of the flow
        //driver.switchTo().defaultContent();
         String parentWindow = driver.getWindowHandle();


        wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//a[span[text()='PDF']]"))).click();
        switchToNewWindow(driver);

        new WebDriverWait(driver, Duration.ofSeconds(10))
            .until(d -> d.getWindowHandles().size() > 1);

        // Switch to the new window
        for (String win : driver.getWindowHandles()) {
            if (!win.equals(parentWindow)) {
                driver.switchTo().window(win);
                break;
            }
        }

        wait.until(ExpectedConditions.elementToBeClickable(
                By.id("printTabs_2_center"))).click();
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("subsetIdSelect")));
        new Select(driver.findElement(By.id("subsetIdSelect"))).selectByVisibleText("IDT");
        WebElement button = driver.findElement(By.xpath("//div[@id='printTabs_2']//button[.//span[normalize-space(.)='Finish & PDF']]"));
        ((JavascriptExecutor)driver).executeScript("arguments[0].click();", button);

        Thread.sleep(6000);

        //wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(@class,'standard_button')]"))).click();
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(@class,'standard_button') and @id='PDFButton']"))).click();
       // switchToNewWindow(driver);

        // wait.until(ExpectedConditions.textToBePresentInElementLocated(By.id("processtextinner"), "Done"));
         long startTime = System.currentTimeMillis();
        // wait.until(ExpectedConditions.elementToBeClickable(By.linkText("click here"))).click();

        Path downloaded = waitForLatestPdf(downloadsDir, startTime);

        // ====== ADDED: save file as <AccountNumber>.pdf for each iteration ======
        Path renamed = downloadsDir.resolve(accountNumber + ".pdf");
        if (Files.exists(renamed)) {
            // avoid collision if file already exists (e.g., rerun)
            renamed = downloadsDir.resolve(accountNumber + "_" + System.currentTimeMillis() + ".pdf");
        }
        Files.move(downloaded, renamed);
        System.out.println("Downloaded: " + downloaded);
        System.out.println("Renamed: " + renamed);
        System.out.println("Done for account: " + accountNumber);

        // ====== ADDED: clean up extra windows and return to root for next iteration ======
        closeAllBut(driver, rootWindow);
        driver.switchTo().window(rootWindow);
        driver.switchTo().defaultContent();
    }

    /**
     * New: using DataFormatter to preserve Excel displayed text and avoid ".0"
     * This is the overloaded method that reads a specific cell (0-indexed row/col).
     * Example: B4 => rowIndex=3, colIndex=1
     */
    private static String readAccountNumber(String excelPath, int rowIndex, int colIndex) throws IOException {
        try (FileInputStream fis = new FileInputStream(excelPath);
                Workbook workbook = WorkbookFactory.create(fis)) {
            Sheet sheet = workbook.getSheet("IPO");
            if (sheet == null) {
                throw new RuntimeException("Sheet 'IPO' not found in " + excelPath);
            }
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                throw new RuntimeException("Missing row " + (rowIndex + 1) + " in sheet IPO");
            }
            Cell cell = row.getCell(colIndex);
            if (cell == null) {
                throw new RuntimeException("Missing cell at row " + (rowIndex + 1) + ", col " + (colIndex + 1));
            }

            DataFormatter formatter = new DataFormatter();
            String raw = formatter.formatCellValue(cell);
            if (raw == null) {
                throw new RuntimeException("Account number cell is empty (row " + (rowIndex + 1) + ", col " + (colIndex + 1) + ")");
            }
            raw = raw.trim();

            // Remove trailing ".0" for integer-like numeric values (1001.0 -> 1001)
            if (raw.matches("^-?\\d+\\.0+$")) {
                raw = raw.replaceAll("\\.0+$", "");
            }
            return raw;
        }
    }

    /**
     * Original single-parameter readAccountNumber left for backward compatibility:
     * reads A2 (rowIndex=1, colIndex=0)
     */
    private static String readAccountNumber(String excelPath) throws IOException {
        return readAccountNumber(excelPath, 1, 0);
    }

    private static void switchToNewWindow(WebDriver driver) {
        String current = driver.getWindowHandle();
        for (String handle : driver.getWindowHandles()) {
            if (!handle.equals(current)) {
                driver.switchTo().window(handle);
            }
        }
    }

    private static Path waitForLatestPdf(Path downloadsDir, long startTime) throws InterruptedException, IOException {
        int attempts = 0;
        while (attempts++ < 60) {
            Optional<Path> latest = Files.list(downloadsDir)
                    .filter(p -> p.toString().toLowerCase().endsWith(".pdf"))
                    .filter(p -> lastModifiedMillis(p) > startTime)
                    .max(Comparator.comparingLong(PDFautomation::lastModifiedMillis));
            if (latest.isPresent()) {
                return latest.get();
            }
            Thread.sleep(1000);
        }
        throw new RuntimeException("PDF download not detected");
    }

    private static long lastModifiedMillis(Path p) {
        try {
            return Files.getLastModifiedTime(p).toMillis();
        } catch (IOException e) {
            return 0L;
        }
    }

    // ====== UPDATED: helper to read up to N account numbers from Excel column (B2 downward) ======
    private static List<String> readAccountNumbers(String excelPath, int startRowIndex, int colIndex, int maxCount) throws IOException {
        // startRowIndex is 0-based; B2 == 1; colIndex for column B is 1.
        List<String> list = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(excelPath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            DataFormatter formatter = new DataFormatter();
            Sheet sheet = workbook.getSheet("IPO");
            if (sheet == null) {
                throw new RuntimeException("Sheet 'IPO' not found in " + excelPath);
            }

            // Ensure we begin from the first data row (header + 1), i.e., B2
            int headerRow = sheet.getFirstRowNum();          // usually 0 for the header
            int dataStart = headerRow + 1;                   // row right after header
            int start = Math.max(startRowIndex, dataStart);  // enforce starting at B2 or beyond

            int last = sheet.getLastRowNum();
            for (int r = start; r <= last && list.size() < maxCount; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell cell = row.getCell(colIndex);
                if (cell == null) continue;

                String raw = formatter.formatCellValue(cell);
                if (raw == null) continue;
                raw = raw.trim();
                if (raw.isEmpty()) continue;

                // Skip common header texts, just in case
                String normalized = raw.replaceAll("\\s+", "").toLowerCase();
                if (normalized.equals("account") || normalized.equals("accountnumber")) {
                    continue;
                }

                // Accept only numeric-like values; normalize 1001.0 -> 1001
                if (raw.matches("^\\d+(?:\\.0+)?$")) {
                    raw = raw.replaceAll("\\.0+$", "");
                    list.add(raw);
                } else {
                    // Non-numeric in the account column: skip
                    continue;
                }
            }
        }
        // De-duplicate while preserving order
        Set<String> dedup = new LinkedHashSet<>(list);
        return new ArrayList<>(dedup);
    }

    // ====== ADDED: close all windows except the given one (root), to keep each iteration clean ======
    private static void closeAllBut(WebDriver driver, String keepHandle) {
        for (String h : new ArrayList<>(driver.getWindowHandles())) {
            if (!h.equals(keepHandle)) {
                try {
                    driver.switchTo().window(h);
                    driver.close();
                } catch (Exception ignore) {}
            }
        }
        try {
            driver.switchTo().window(keepHandle);
        } catch (Exception ignore) {}
    }
}