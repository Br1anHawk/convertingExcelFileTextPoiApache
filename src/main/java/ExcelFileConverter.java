import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;

public class ExcelFileConverter {


    public void convert(ArrayList<String> filePaths) throws IOException {
        String sourceFolderPath = filePaths.get(0).substring(0, filePaths.get(0).lastIndexOf("/") + 1);
        String destinationFilePathXLS = sourceFolderPath + "combinedFile.xls";
        ArrayList<Sheet> sheets = initSheets(filePaths);
        OutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(destinationFilePathXLS);
            Workbook workbook = combineContentsFromAllSheets(sheets);
            workbook.write(fileOutputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            assert fileOutputStream != null;
            fileOutputStream.close();
        }
        ToCSV converterCSV = new ToCSV();
        converterCSV.convertExcelToCSV(destinationFilePathXLS, sourceFolderPath);
    }

    private ArrayList<Sheet> initSheets(ArrayList<String> filePaths) throws IOException {
        ArrayList<Sheet> sheets = new ArrayList<>();
        InputStream fileInputStream = null;
        try {
            for (String filePath : filePaths) {
                fileInputStream = new FileInputStream(filePath);
                Workbook workbook = WorkbookFactory.create(fileInputStream);
                Sheet sheet = workbook.getSheetAt(0);
                sheets.add(sheet);
                fileInputStream.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            assert fileInputStream != null;
            fileInputStream.close();
            return sheets;
        }
    }

    private Workbook combineContentsFromAllSheets(ArrayList<Sheet> sheets) {
        Workbook workbook = new HSSFWorkbook();
        Sheet combinedContentSheet = workbook.createSheet("sheet");
        int rowNumber = 0;
        Row row = combinedContentSheet.createRow(rowNumber);
        row.createCell(0).setCellValue("Имя точки");
        row.createCell(1).setCellValue("Код точки");
        row.createCell(2).setCellValue("Серийный номер");
        row.createCell(3).setCellValue("Правило применения коэффициентов (устройство = УСПД)");
        row.createCell(4).setCellValue("KI");
        row.createCell(5).setCellValue("KU");
        row.createCell(6).setCellValue("Дата применения коэффициентов");
        row.createCell(7).setCellValue("Тип прибора учета");
        row.createCell(8).setCellValue("Дата установки прибора");
        row.createCell(9).setCellValue("Номер порта");
        rowNumber++;
        int portNumber = 0;
        for (Sheet sheet : sheets) {
            portNumber++;
            int contentRangeBeginPosition = 1;
            while (sheet.getRow(contentRangeBeginPosition).getCell(0).toString().isEmpty()) contentRangeBeginPosition++;
            contentRangeBeginPosition += 2;
            while (sheet.getRow(contentRangeBeginPosition) != null) {
                String meterType = sheet.getRow(contentRangeBeginPosition).getCell(1).getStringCellValue();
                String meterNumber = "0" + (int) sheet.getRow(contentRangeBeginPosition).getCell(2).getNumericCellValue();
                String meterConnectionPoint = modifyMeterConnectionPoint(sheet.getRow(contentRangeBeginPosition).getCell(3).getStringCellValue());
                String meterKI = String.valueOf((int) sheet.getRow(contentRangeBeginPosition).getCell(4).getNumericCellValue());
                if (meterKI.equals("0")) meterKI = "1";
                String meterKU = "1";
                String KIRule = "Данные в устройстве без коэффициентов";
                String dateOfInstallationKI = "01.01.2021";
                String dateOfInstallationMeter = "01.01.2021";

                row = combinedContentSheet.createRow(rowNumber);
                row.createCell(0).setCellValue(meterConnectionPoint);
                row.createCell(1).setCellValue(meterConnectionPoint);
                row.createCell(2).setCellValue(meterNumber);
                row.createCell(3).setCellValue(KIRule);
                row.createCell(4).setCellValue(meterKI);
                row.createCell(5).setCellValue(meterKU);
                row.createCell(6).setCellValue(dateOfInstallationKI);
                row.createCell(7).setCellValue(meterType);
                row.createCell(8).setCellValue(dateOfInstallationMeter);
                row.createCell(9).setCellValue(portNumber);

                rowNumber++;
                contentRangeBeginPosition++;
            }
        }
        return workbook;
    }

    private String modifyMeterConnectionPoint(String meterConnectionPoint) {
        String modified = meterConnectionPoint;
        modified = modified.substring(modified.indexOf("\\") + 1);

        return modified;
    }

}
