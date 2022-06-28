import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class WorkBook {
    private static final String FILE = "C:/Users/79963/Desktop/buh/src/main/resources/main.xls";

    private static final SimpleDateFormat FORMAT = new SimpleDateFormat("MM.yyyy");
    private static final SimpleDateFormat SIMPLE_DATE_FORMAT = new SimpleDateFormat("MM.yy");

    private static final String SERVICE_SHEET_NAME = "Служебка Шали ";
    private static final String LIZING_SHEET_NAME = "лизинг Шали-Агро";

    private static final String DATE_COLUMN_NAME = "Дата";
    private static final String CONTRACT_NUMBER_COLUMN_NAME = "Номер договора";
    private static final String SUM_COLUMN_NAME = "Сумма платежа";
    private static final String LIZING_COLUMN_NAME = "Лизингодатель";
    private static final String TECHNIQUE_COLUMN_NAME = "Техника";
    private static final String DIRECTION_COLUMN_NAME = "Направление";

    private static final String DATE_MATCHER = "[0-9]{2}\\.[0-9]{2}\\.[0-9]{4}";
    private static final String EMPTY = "";

    public static void addStandardWorkBook() throws IOException {
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        hssfWorkbook.write(new FileOutputStream(FILE));
        hssfWorkbook.close();
    }

    public static void addServiceSheet() throws IOException {
        Date date = new Date();
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(FILE));
        Sheet sheet = wb.createSheet(SERVICE_SHEET_NAME + SIMPLE_DATE_FORMAT.format(date));
        Row row = sheet.createRow(0);

        Cell firstCell = row.createCell(0);
        firstCell.setCellValue(DATE_COLUMN_NAME);

        Cell secondCell = row.createCell(1);
        secondCell.setCellValue(CONTRACT_NUMBER_COLUMN_NAME);

        Cell thirdCell = row.createCell(2);
        thirdCell.setCellValue(SUM_COLUMN_NAME);

        Cell fourthCell = row.createCell(3);
        fourthCell.setCellValue(LIZING_COLUMN_NAME);

        Cell fifthCell = row.createCell(4);
        fifthCell.setCellValue(TECHNIQUE_COLUMN_NAME);

        Cell sixthCell = row.createCell(5);
        sixthCell.setCellValue(DIRECTION_COLUMN_NAME);

        wb.write(new FileOutputStream(FILE));
        wb.close();
    }

    public static void generateNewService() throws IOException {
        Date date = new Date();
        String dateNow = FORMAT.format(date);
        int index = 13;
        int thirdRowIndex = 1;
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(FILE));
        Sheet firstSheet = wb.getSheet(LIZING_SHEET_NAME);
        Sheet secondSheet = wb.getSheet(SERVICE_SHEET_NAME + SIMPLE_DATE_FORMAT.format(date));

        for (int i = 0; i < 38; i++) {
            String contractName = EMPTY;
            String contractDate = EMPTY;
            double contractPrice = 0;
            String lizingName = EMPTY;
            String transportName = EMPTY;

            for (int j = 0; j < 62; j++) {
                Row firstRow = firstSheet.getRow(j);
                Cell firstCell = firstRow.getCell(index);
                String contract = firstCell.getStringCellValue();
                if (contract.matches(DATE_MATCHER)) {
                    if (contract.substring(3, 10).equals(dateNow)) {
                        contractDate = contract;
                        contractPrice = firstRow.getCell(index + 1).getNumericCellValue();

                        for (int k = 1; k < 40; k++) {
                            Row secondRow = firstSheet.getRow(k);
                            Cell secondCell = secondRow.getCell(0);
                            if (secondCell.getStringCellValue().equals(contractName)) {
                                lizingName = secondRow.getCell(4).getStringCellValue();
                                transportName = secondRow.getCell(6).getStringCellValue();
                            }
                        }

                    }
                } else if (!contract.matches(DATE_MATCHER) && !contract.equals(EMPTY)){
                    contractName = contract;
                }
            }

            Row thirdRow = secondSheet.createRow(thirdRowIndex);

            Cell dateCell = thirdRow.createCell(0);
            dateCell.setCellValue(contractDate);

            Cell nameCell = thirdRow.createCell(1);
            nameCell.setCellValue(contractName);

            Cell priceCell = thirdRow.createCell(2);
            priceCell.setCellValue(contractPrice);

            Cell lizingCell = thirdRow.createCell(3);
            lizingCell.setCellValue(lizingName);

            Cell transportCell = thirdRow.createCell(4);
            transportCell.setCellValue(transportName);

            index += 3;
            thirdRowIndex++;
        }
        wb.write(new FileOutputStream(FILE));
        wb.close();
    }
}
