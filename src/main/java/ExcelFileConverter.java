import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;

public class ExcelFileConverter {

    public boolean isConvertedFileExists(ArrayList<File> files, String routerID) {
        String sourceFolderPath = files.get(0).getPath().substring(0, files.get(0).getPath().lastIndexOf("\\") + 1);
        String destinationFilePathXLS = sourceFolderPath + "Router_" + routerID + ".xls";
        File file = new File(destinationFilePathXLS);
        if (file.exists()) {
            return true;
        } else {
            return false;
        }
    }

    public void convert(ArrayList<File> files, String routerID) throws IOException {
        String sourceFolderPath = files.get(0).getPath().substring(0, files.get(0).getPath().lastIndexOf("\\") + 1);
        String destinationFilePathXLS = sourceFolderPath + "Router_" + routerID + ".xls";
        String destinationFilePathPortsInfoXLS = sourceFolderPath + "Router_" + routerID + "_ports_info.xls";
        ArrayList<Sheet> sheets = initSheets(files);
        OutputStream fileOutputStream = null;
        try {
            fileOutputStream = new FileOutputStream(destinationFilePathXLS);
            Workbook workbook = combineContentsFromAllSheets(sheets);
            workbook.write(fileOutputStream);
            fileOutputStream.close();

            fileOutputStream = new FileOutputStream(destinationFilePathPortsInfoXLS);
            workbook = getPortsInfoFromAllSheets(sheets);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            assert fileOutputStream != null;
            fileOutputStream.close();
        }
        ToCSV converterCSV = new ToCSV();
        converterCSV.convertExcelToCSV(destinationFilePathXLS, sourceFolderPath);
    }

    private ArrayList<Sheet> initSheets(ArrayList<File> files) throws IOException {
        ArrayList<Sheet> sheets = new ArrayList<>();
        InputStream fileInputStream = null;
        try {
            for (File file : files) {
                fileInputStream = new FileInputStream(file);
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
            while (sheet.getRow(contentRangeBeginPosition) == null) contentRangeBeginPosition++;
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
        String pointConnection = modified.substring(modified.lastIndexOf("\\") + 1);
        if (String.valueOf(pointConnection.charAt(0)).matches("[0-9]")) {
            pointConnection = "кв. " + pointConnection;
        }
        return pointConnection;
    }

    private Workbook getPortsInfoFromAllSheets(ArrayList<Sheet> sheets) {
        Workbook workbook = new HSSFWorkbook();
        Sheet combinedContentSheet = workbook.createSheet("ports_info");
        int rowNumber = 0;
        Row row = combinedContentSheet.createRow(rowNumber);
        row.createCell(0).setCellValue("PortNumber");
        row.createCell(1).setCellValue("BuildingAddress");
        rowNumber++;
        int portNumber = 0;
        for (Sheet sheet : sheets) {
            portNumber++;
            int contentRangeBeginPosition = 1;
            while (sheet.getRow(contentRangeBeginPosition) == null) contentRangeBeginPosition++;
            contentRangeBeginPosition += 2;
            String meterConnectionPointFull = getAddress(sheet.getRow(contentRangeBeginPosition).getCell(3).getStringCellValue());
            row = combinedContentSheet.createRow(rowNumber);
            row.createCell(0).setCellValue(portNumber);
            row.createCell(1).setCellValue(meterConnectionPointFull);
            rowNumber++;
        }
        return workbook;
    }

    private String getAddress(String meterConnectionPointFull) {
        int indexPosition = meterConnectionPointFull.indexOf("\\п ");
        if (indexPosition == -1) {
            indexPosition = meterConnectionPointFull.indexOf("\\эт ");
        }
        if (indexPosition == -1) {
            indexPosition = meterConnectionPointFull.lastIndexOf("\\");
        }
        meterConnectionPointFull = meterConnectionPointFull.substring(0, indexPosition);
        return meterConnectionPointFull;
    }
}
