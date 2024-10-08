import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
public class MainClass {
    public static void main(String[] args) {
        // TODO Auto-generated method stub
        Map<Integer, Integer> File = new HashMap<Integer, Integer>();
        String path = System.getProperty("user.dir");

        String[] WAOC = new String[3];
        String[] WAPO = new String[3];

        WAOC[0] = "C:\\Users\\2174112\\Downloads\\WAOC00826050624130548197.xml";
        WAOC[1] = "C:\\Users\\2174112\\Downloads\\WAOC84096050624130548114.xml";
        WAOC[2] = "C:\\Users\\2174112\\Downloads\\WAOC01427050624130548342.xml";
//        // WAOC[3] = "C:\\Users\\919976\\Downloads\\WAOC84096121023114520366.xml";
//
        WAPO[0] = "C:\\Users\\2174112\\Downloads\\WAPO01427050624130547937.xml";
        WAPO[1] = "C:\\Users\\2174112\\Downloads\\WAPO84096050624130547689.xml";
        WAPO[2] = "C:\\Users\\2174112\\Downloads\\WAPO00826050624130547826.xml";
//        // WAPO[3] = "C:\\Users\\919976\\Downloads\\WAPO84096121023114519782.xml";

//
//		 File = testXMLResponseusingfilename(WAOC);
//		 String excelFilePath = "C:\\Users\\2174112\\Downloads\\WAOCDataSheet.xlsx";



        File = testXMLResponseusingfilename(WAPO);
        String excelFilePath = "C:\\Users\\2174112\\Downloads\\WAPODataSheet.xlsx";


        System.out.println("All Files Content with size :" + File.size());
        System.out.println("-------------------------------------------------");
        // Print the map
        for (Map.Entry<Integer, Integer> entry : File.entrySet()) {

            System.out.println(entry.getKey() + " " + entry.getValue());
        }
        System.out.println("-------------------------------------------------");
        System.out.println(" ");

        createexcel(File);

        String sheetname = "TestData";
        String columnname1 = "ProdCode";
        String columnname2 = "Quantity";

        Map<String, String> orderMap = readExcelData(excelFilePath, sheetname, columnname1, columnname2);
        System.out.println("Expected Map : " + orderMap);
        Map<Integer, Integer> integerMap = convertMap(orderMap);

        // Actual and Expected
        boolean result = compareMaps(File, integerMap);
        if (result) {
            System.out.println("Maps are equal.");
        } else {
            System.out.println("Maps are not equal.");
        }

    }

    public static Map<Integer, Integer> testXMLResponseusingfilename(String[] arr) {

        Map<Integer, Integer> itemQuantityMap = new HashMap<Integer, Integer>();
        int n = 0;

        for (int k = 0; k < arr.length; k++) {

            String filePath = arr[k];
            // Read XML file
            try {
                String xmlContent = new String(Files.readAllBytes(Paths.get(filePath)));

                // Parse XML content
                DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
                DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
                Document doc = dBuilder.parse(new File(filePath));
                doc.getDocumentElement().normalize();

                // Get the NodeList for OrderLine elements
                NodeList orderLineList = doc.getElementsByTagName("OrderLine");

                for (int i = 0; i < orderLineList.getLength(); i++) {
                    Element orderLineElement = (Element) orderLineList.item(i);

                    // Get CustItem code
                    int custItemCode = Integer
                            .parseInt(orderLineElement.getElementsByTagName("Code").item(0).getTextContent());

                    // Get OrderQty units
                    NodeList orderQtyList = orderLineElement.getElementsByTagName("OrderQty");
                    int orderQty = 0;
                    if (orderQtyList.getLength() > 0) {
                        Element orderQtyElement = (Element) orderQtyList.item(0);
                        orderQty = Integer
                                .parseInt(orderQtyElement.getElementsByTagName("Unit").item(0).getTextContent());
                    }

                    // Store the values in the map
                    itemQuantityMap.put(custItemCode, orderQty);
                }

                // Print the map
                System.out.println(
                        "After File Path " + filePath + "Reading toatl Content with size : " + itemQuantityMap.size());
                System.out.println("-------------------------------------------------");
                for (Map.Entry<Integer, Integer> entry : itemQuantityMap.entrySet()) {
                    // System.out.println("CustItem Code " + entry.getKey() + ", OrderQty Units: " +
                    // entry.getValue());
                }
                System.out.println("-------------------------------------------------");
                System.out.println(" ");

            }

            catch (Exception e) {
                e.printStackTrace();
            }

        }

        return itemQuantityMap;
    }

    public static void createexcel(Map<Integer, Integer> map) {
        // Create a new Excel workbook
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Map Data");

            // Create the header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("ProdCode");
            headerRow.createCell(1).setCellValue("Quantity");

            int rowNumber = 1;
            for (Map.Entry<Integer, Integer> entry : map.entrySet()) {
                Row row = sheet.createRow(rowNumber++);
                row.createCell(0).setCellValue(entry.getKey());
                row.createCell(1).setCellValue(entry.getValue());
            }

            String filepath = "C:\\Users\\2174112\\Downloads\\DataFile.xlsx";

            // Write the workbook content to a file
            try (FileOutputStream fileOut = new FileOutputStream(filepath)) {
                workbook.write(fileOut);
                System.out.println("Excel file has been created successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // By passing excelFilePath,sheetname,custIdColumnName,qtyColumnName

    public static Map<String, String> readExcelData(String excelFilePath, String sheetname, String custIdColumnName,
                                                    String qtyColumnName) {
        Map<String, String> orderMap = new HashMap<>();

        try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
             Workbook workbook = WorkbookFactory.create(fileInputStream)) {
            Sheet sheet = workbook.getSheet(sheetname);

            if (sheet == null) {
                System.out.println("Sheet with name " + sheetname + " not found");
                return orderMap; // Return an empty map if sheet not found
            }

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                System.out.println("Header row not found in the sheet.");
                return orderMap; // Return an empty map if header row not found
            }

            int custIdIndex = -1;
            int qtyIndex = -1;

            Iterator<Cell> cellIterator = headerRow.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String columnHeader = cell.getStringCellValue();
                if (columnHeader.equalsIgnoreCase(custIdColumnName)) {
                    custIdIndex = cell.getColumnIndex();
                } else if (columnHeader.equalsIgnoreCase(qtyColumnName)) {
                    qtyIndex = cell.getColumnIndex();
                }
            }

            if (custIdIndex == -1 || qtyIndex == -1) {
                System.out.println("Required columns missing in the sheet.");
                return orderMap; // Return an empty map if required columns are missing
            }

            // Read data from subsequent rows
            Iterator<Row> rowIterator = sheet.iterator();
            // Skip the header row
            rowIterator.next();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                if (row.getCell(custIdIndex) != null && row.getCell(qtyIndex) != null) { // Add null check here
                    String custId;
                    String qty;
                    if (row.getCell(custIdIndex).getCellType() == CellType.STRING) {
                        custId = row.getCell(custIdIndex).getStringCellValue();
                        qty = row.getCell(qtyIndex).getStringCellValue();

                    } else {
                        custId = String.valueOf((int) row.getCell(custIdIndex).getNumericCellValue());
                        qty = String.valueOf((int) row.getCell(qtyIndex).getNumericCellValue());
                    }
                    orderMap.put(custId, qty);
                }
            }

        } catch (IOException | EncryptedDocumentException ex) {
            ex.printStackTrace();
        }

        return orderMap;
    }

    // Compare maps
    public static boolean compareMaps(Map<Integer, Integer> map1, Map<Integer, Integer> map2) {
        boolean equal = true;

        // Check keys in map1 but not in map2
        for (Integer key : map1.keySet()) {
            if (!map2.containsKey(key)) {
                System.out.println("Key in Actual but not in Expected: " + key);
                equal = false;
            }
        }

        // Check keys in map2 but not in map1
        for (Integer key : map2.keySet()) {
            if (!map1.containsKey(key)) {
                System.out.println("Key in Expected but not in Actual: " + key);
                equal = false;
            }
        }

        // Check if corresponding values of keys in map1 and map2 are different
        for (Integer key : map1.keySet()) {
            if (map2.containsKey(key) && !map1.get(key).equals(map2.get(key))) {
                System.out.println("Key with different values: " + key + " - Actual value: " + map1.get(key)
                        + ", Expected value: " + map2.get(key));
                equal = false;
            }
        }

        return equal;
    }

    public static Map<Integer, Integer> convertMap(Map<String, String> stringMap) {
        Map<Integer, Integer> integerMap = new HashMap<>();

        // Iterate through the entries of the stringMap
        for (Map.Entry<String, String> entry : stringMap.entrySet()) {
            try {
                // Convert the keys and values from String to Integer
                Integer key = Integer.parseInt(entry.getKey());
                Integer value = Integer.parseInt(entry.getValue());

                // Put the converted key-value pair into the integerMap
                integerMap.put(key, value);
            } catch (NumberFormatException e) {
                // Handle the case where parsing fails
                System.err.println("Failed to parse integer: " + e.getMessage());
            }
        }

        return integerMap;
    }
}
