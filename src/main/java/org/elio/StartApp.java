package org.elio;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.BaseFormulaEvaluator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * created by elio on 12/09/2022
 */
public class StartApp {
    private static final SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
    private static final SimpleDateFormat sdfForFileName = new SimpleDateFormat("yyyyMMdd");
    private static final SimpleDateFormat sdfForAccSheet = new SimpleDateFormat("yyyy/MM/dd");
    private static final SimpleDateFormat sdfForInvoice = new SimpleDateFormat("yyyy-MM-dd");
    private static final Calendar instance = Calendar.getInstance();
    private static final Scanner sc = new Scanner(System.in);
    private static ConfigWrapper config;

    private static String IGNORE_COL = "";
    private static final int[] endSheetAndCol = new int[]{99, 99};
    private static Date START_DATE;
    private static Date END_DATE;
    private static final HashMap<String, BigDecimal> INVOICE_MAP = new HashMap<>();

    static {
        try {
            config = new ConfigWrapper();
            populateInvoiceMap(INVOICE_MAP);
        } catch (IllegalAccessException e) {
            System.exit(-1);
        } catch (IOException | ParseException e) {
            throw new RuntimeException(e);
        }
    }

    public static void main(String[] args) {

        // initiate Calendar with user input
        int month = instance.get(Calendar.MONTH);
        System.out.print("输入要填写的报销月份(" + (month + 1) + ")月:");
        String s = sc.nextLine();
        int startMonth = "".equals(s) ? month : Integer.parseInt(s) - 1;
        instance.set(Calendar.MONTH, startMonth);
        instance.set(Calendar.DAY_OF_MONTH, config.START_DAY_OF_MONTH);
        START_DATE = instance.getTime();
        // find out col num that do not need to process
        findOutSatAndSun();
        // get template File with user input
        File templateFile = getExcelFile();

        try (FileInputStream is = new FileInputStream(templateFile);
             Workbook workbook = WorkbookFactory.create(is)) {
            for (int i = 0; i < 5; i++) {
                clearForm(workbook, i);
                generateTime(workbook, i);
                fillForm(workbook, i);
            }
            fillAccountMonthlySheet(workbook);
            is.close();
            int dotIndex = templateFile.getName().lastIndexOf('.');
            String newFileName = config.TEMPLATE_FOLDER_PATH +
                    File.separator +
                    "报销单(" +
                    sdfForFileName.format(START_DATE) +
                    "-" +
                    sdfForFileName.format(END_DATE) +
                    "-" +
                    config.NAME +
                    ")" +
                    templateFile.getName().substring(dotIndex);
            workbook.write(Files.newOutputStream(new File(newFileName).toPath()));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void populateInvoiceMap(HashMap<String, BigDecimal> invoiceMap) throws IOException, ParseException {

        File[] routeFiles = new File(config.TAXI_INVOICE_PATH).listFiles(filename -> filename.getName().endsWith(".pdf") && filename.getName().contains("行程单"));
        if (routeFiles == null) {
            System.out.println("No routine PDF File");
            return;
        }
        PDFTextStripper stripper = new PDFTextStripper();
        Pattern r1 = Pattern.compile("-(.*?)元-");
        Pattern r2 = Pattern.compile("行程时间：(.*?) [0-9]+:");
        for (File file : routeFiles) {
            Matcher m1 = r1.matcher(file.getName());
            if (!m1.find()) {
                continue;
            }
            BigDecimal price = new BigDecimal(m1.group(1));
            PDDocument document = PDDocument.load(file);
            Matcher m2 = r2.matcher(stripper.getText(document));
            document.close();
            if (!m2.find()) {
                continue;
            }
            String date = m2.group(1);
            if (invoiceMap.containsKey(date)) {
                price = invoiceMap.get(date).add(price);
            }
            invoiceMap.put(date, price);
        }
        System.out.println(invoiceMap);


    }

    private static File getExcelFile() {
        File folder = new File(config.TEMPLATE_FOLDER_PATH);
        File[] files = folder.listFiles(pathname -> pathname.getName().endsWith(".xlsx") || pathname.getName().endsWith(".xls"));
        if (files == null || files.length == 0) {
            System.exit(-1);
        }
        System.out.println("-----------------------");
        for (int i = 0; i < files.length; i++) {
            System.out.println("[" + i + "]" + files[i].getName());
        }
        System.out.print("请选择修改的模板文件(0):");
        String fileSeq = sc.nextLine();
        return files["".equals(fileSeq) ? 0 : Integer.parseInt(fileSeq)];
    }

    private static void generateTime(Workbook workbook, int i) throws ParseException {
        for (int j = config.CONTENT_START_COL; j < config.CONTENT_START_COL + 7; j++) {
            if (instance.get(Calendar.DAY_OF_MONTH) == config.END_DAY_OF_MONTH + 1 && i != 0) {
                endSheetAndCol[0] = i;
                endSheetAndCol[1] = j - 1;
                break;
            }
            Cell cell = workbook.getSheetAt(i).getRow(2).getCell(j);
            // remove h m s
            cell.setCellValue(sdf.parse(sdf.format(instance.getTime())));
            instance.add(Calendar.DAY_OF_MONTH, 1);
        }
    }

    private static void clearForm(Workbook workbook, int i) {
        Sheet sheet = workbook.getSheetAt(i);
        // initiate(clear and create) cells
        for (int j = config.CONTENT_START_COL; j < config.CONTENT_START_COL + 7; j++) {
            for (int k = config.CONTENT_START_ROW - 1; k < config.CONTENT_END_ROW + 1; k++) {
                if (k == config.HEADER_IGNORE_ROW) {
                    continue;
                }
                Cell cell = sheet.getRow(k).getCell(j);
                if (cell == null) {
                    cell = sheet.getRow(k).createCell(i);
                }
                cell.setCellValue("");
            }
        }
    }

    private static void fillForm(Workbook workbook, int i) {
        if (endSheetAndCol[0] < i) {
            return;
        }
        Sheet sheet = workbook.getSheetAt(i);
        // fill the form
        for (int j = config.CONTENT_START_COL; j < config.CONTENT_START_COL + 7; j++) {
            if (endSheetAndCol[0] == i && endSheetAndCol[1] < j) {
                break;
            }
            if (IGNORE_COL.contains(String.valueOf(j))) {
                continue;
            }
            for (int k = config.CONTENT_START_ROW; k < config.CONTENT_END_ROW + 1; k++) {
                if (k == config.HEADER_IGNORE_ROW) {
                    continue;
                }
                Cell cell = sheet.getRow(k).getCell(j);
                if (cell == null) {
                    cell = sheet.getRow(k).createCell(i);
                }
                if (k == config.TRAIN_ROW) {
                    cell.setCellValue(config.FEE_OF_TRAIN);
                } else if (k == config.FOOD_ROW) {
                    cell.setCellValue(config.FEE_OF_FOOD);
                } else if (k == config.BUSINESS_TRIP_ROW) {
                    cell.setCellValue(config.FEE_OF_BUSINESS_TRIP);
                } else if (k == config.PROJECT_LOCATION_ROW) {
                    cell.setCellValue(config.PROJECT_LOCATION);
                } else if (k == config.PROJECT_NUMBER_ROW) {
                    cell.setCellValue(config.PROJECT_NUMBER);
                } else if (k == config.TAXI_ROW) {
                    Date date = sheet.getRow(2).getCell(j).getDateCellValue();
                    if (INVOICE_MAP.get(sdfForInvoice.format(date)) != null) {
                        cell.setCellValue(INVOICE_MAP.get(sdfForInvoice.format(date)).doubleValue());
                    }
                }
            }
        }
        // update vertical accumulate formula
        for (int j = config.ACCUMULATE_VER_ROW; j < sheet.getLastRowNum() + 1; j++) {
            Row row = sheet.getRow(j);
            if (row == null) {
                continue;
            }
            Cell cell = row.getCell(config.ACCUMULATE_VER_COL);
            updateFormula(workbook, cell);
        }

        // update horizontal accumulate formula
        Row row = sheet.getRow(config.ACCUMULATE_HOR_ROW);
        for (int col = config.ACCUMULATE_HOR_COL; col < row.getLastCellNum() + 1; col++) {
            Cell cell = row.getCell(col);
            updateFormula(workbook, cell);
        }

        // fill name and department
        Cell nameCell = sheet.getRow(0).getCell(10);
        Cell staffCell = sheet.getRow(33).getCell(10);
        Cell deptCell = sheet.getRow(1).getCell(10);
        nameCell.setCellValue(config.NAME);
        staffCell.setCellValue(config.NAME);
        deptCell.setCellValue(config.DEPARTMENT);
    }

    private static void fillAccountMonthlySheet(Workbook workbook) throws ParseException {
        Sheet sheet = workbook.getSheetAt(config.ACCUMULATE_SHEET);
        // update accumulate sheet formula
        for (int col = 1; col < 4; col++) {
            Row sheetRow = sheet.getRow(col);
            for (int row = 4; row < 7; row++) {
                Cell cell = sheetRow.getCell(row);
                updateFormula(workbook, cell);
            }
        }
        Cell dateCell = sheet.getRow(1).getCell(1);
        dateCell.setCellValue(sdf.parse(sdf.format(instance.getTime())));
        Cell timeCell = sheet.getRow(4).getCell(0);
        instance.add(Calendar.DAY_OF_MONTH, -1);
        END_DATE = instance.getTime();
        timeCell.setCellValue(sdfForAccSheet.format(START_DATE) + "-" + sdfForAccSheet.format(END_DATE));

        // fill name
        sheet.getRow(0).createCell(1).setCellValue(config.NAME);
    }

    private static void findOutSatAndSun() {
        int col = config.CONTENT_START_COL;
        for (int dayOfWeek = instance.get(Calendar.DAY_OF_WEEK); dayOfWeek < Calendar.SATURDAY + 1; dayOfWeek++, col++) {
            if (dayOfWeek == Calendar.SATURDAY || dayOfWeek == Calendar.SUNDAY) {
                IGNORE_COL += "/" + col;
            }
        }
        if (IGNORE_COL.lastIndexOf("/") == 0) {
            IGNORE_COL += "/" + col;
        }
    }

    private static void updateFormula(Workbook wb, Cell cell) {
        if (cell == null) {
            return;
        }
        BaseFormulaEvaluator eval = null;
        if (wb instanceof HSSFWorkbook)
            eval = new HSSFFormulaEvaluator((HSSFWorkbook) wb);
        else if (wb instanceof XSSFWorkbook)
            eval = new XSSFFormulaEvaluator((XSSFWorkbook) wb);
        if (eval == null) {
            System.out.println("error--");
            System.exit(-1);
        }

        if (cell.getCellType() == CellType.FORMULA) {
            eval.evaluateFormulaCell(cell);
        }

    }
}
