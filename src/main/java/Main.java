import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

class PDFBoxExample {
    public static void main(String[] args) throws IOException {
        File folder = new File("C:\\Users\\claytonc\\OneDrive - The Biovac Institute\\Desktop\\TEST MATRIX");
        File[] listOfFiles = folder.listFiles();

        // Create a new Excel workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        // Create a new sheet in the workbook
        XSSFSheet sheet = workbook.createSheet("Referenced Documents");

        int rowNum = 0;
        for (File file : listOfFiles) {
            if (file.isFile() && file.getName().endsWith(".pdf")) {
                PDDocument document = PDDocument.load(file);
                PDFTextStripper textStripper = new PDFTextStripper();

                // Extract the PDF file name
                String fileName = file.getName();

                // Extract the text under the heading "Referenced Documents"
                int pageNum = 1;
                boolean found = false;
                while (pageNum <= document.getNumberOfPages() && !found) {
                    textStripper.setStartPage(pageNum);
                    textStripper.setEndPage(pageNum);
                    String text = textStripper.getText(document);
                    int index = text.indexOf("Referenced Documents");
                    if (index >= 0) {
                        int endIndex = text.indexOf("9. ", index);
                        if (endIndex >= 0) {
                            text = text.substring(index + "Referenced Documents".length(), endIndex);
                        } else {
                            text = text.substring(index + "Referenced Documents".length());
                        }
                        // Split the text by new line and write the extracted information to the Excel sheet
                        String[] referencedDocuments = text.split("\n");
                        Row row = sheet.createRow(rowNum++);
                        row.createCell(0).setCellValue(fileName);
                        for (int i = 0; i < referencedDocuments.length; i++) {
                            row.createCell(i + 1).setCellValue(referencedDocuments[i]);
                        }
                        found = true;
                    }
                    pageNum++;
                }

                document.close();
            }
        }

        // Write the workbook to a file
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\claytonc\\OneDrive - The Biovac Institute\\Desktop\\TEST MATRIX\\ReferencedDocuments.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }
}