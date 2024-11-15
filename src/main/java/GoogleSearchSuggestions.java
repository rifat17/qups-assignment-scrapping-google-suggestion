import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.*;
import java.time.DayOfWeek;
import java.time.Duration;
import java.time.LocalDate;
import java.util.*;


public class GoogleSearchSuggestions {
    private final WebDriver driver;


    public GoogleSearchSuggestions(WebDriver driver) {
        super();
        this.driver = driver;
    }

    public static void main(String[] args) {
        WebDriver driver = new ChromeDriver();
        GoogleSearchSuggestions gss = new GoogleSearchSuggestions(driver);
        try {
            Properties properties = getProperties();
            String filePath = properties.getProperty("file.path");;
//			System.out.println(filePath);

            FileInputStream fis = gss.loadFile(filePath);
            Sheet sheet = gss.loadSheet(fis);
            for (Row row : sheet) {

                if (row.getRowNum() == 0 | row.getRowNum() == 1) {
                    continue; // Skip header rows
                }
//	            System.out.println(row.getRowNum());

                String keyword = row.getCell(2).getStringCellValue();
//	            System.out.println(keyword);
                List<String> suggestions = gss.getGoogleSuggestions(keyword);

                if (suggestions.size() >= 2) {
                    String longestSuggestion = suggestions.get(0);
                    String shortestSuggestion = suggestions.get(1);
//	            	  System.out.println("Longest suggestion: " + longestSuggestion);
//	            	  System.out.println("Shortest suggestion: " + shortestSuggestion);
                    // Update Excel sheet
                    Cell longestSuggestionCell = CellUtil.getCell(row, 3); // Creates cell if missing
                    Cell shortestSuggestionCell = CellUtil.getCell(row, 4); // Creates cell if missing

                    longestSuggestionCell.setCellValue(longestSuggestion);
                    shortestSuggestionCell.setCellValue(shortestSuggestion);
//	            	  row.getCell(3).setCellValue(longestSuggestion);
//	            	  row.getCell(4).setCellValue(shortestSuggestion);
                } else {
                    System.out.println("Fewer than two suggestions found.");
                }

            }
        } catch (IOException | EncryptedDocumentException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } finally {
            driver.close();
        }


    }

    private static Properties getProperties() throws IOException {

        Properties properties = new Properties();
        try (InputStream inputStream = GoogleSearchSuggestions.class.getClassLoader().getResourceAsStream("config.properties")) {
            properties.load(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return properties;
    }

    public Sheet loadSheet(FileInputStream fis) throws EncryptedDocumentException, IOException {
        Workbook workbook = WorkbookFactory.create(fis);

        // Get today's day of the week
        LocalDate today = LocalDate.now();
        DayOfWeek dayOfWeek = today.getDayOfWeek();
        String sheetName = dayOfWeek.toString();


        return workbook.getSheet(sheetName);

    }

    public FileInputStream loadFile(String fileLocation) throws FileNotFoundException {
        FileInputStream fis = new FileInputStream(new File(fileLocation));

        return fis;
    }

    public List<String> getGoogleSuggestions(String q) {

        this.driver.get("https://www.google.com/");

        this.driver.findElement(By.name("q")).sendKeys(q);
        this.driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));

        List<String> suggestions = new ArrayList<>();
        try {
            // Get all suggestions
            List<WebElement> suggestionsElements = this.driver.findElements(By.xpath("//ul[@role='listbox']//li"));

            // Map suggestions to their lengths
            Map<String, Integer> suggestionLengths = new HashMap<>();
            for (WebElement suggestionElement : suggestionsElements) {
                String suggestionText = suggestionElement.getText();
                suggestionLengths.put(suggestionText, suggestionText.length());
            }

            // Find the longest and shortest suggestions
            String longestSuggestion = suggestionLengths.entrySet().stream().max(Comparator.comparingInt(Map.Entry::getValue)).map(Map.Entry::getKey).orElse(null);

            String shortestSuggestion = suggestionLengths.entrySet().stream().min(Comparator.comparingInt(Map.Entry::getValue)).map(Map.Entry::getKey).orElse(null);

            // Add to the list
            suggestions.add(longestSuggestion);
            suggestions.add(shortestSuggestion);

        } catch (Exception e) {
            e.printStackTrace();
        }

        return suggestions;
    }

}
