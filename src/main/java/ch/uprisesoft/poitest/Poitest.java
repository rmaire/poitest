/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Project/Maven2/JavaApp/src/main/java/${packagePath}/${mainClassName}.java to edit this template
 */
package ch.uprisesoft.poitest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author rma
 */
public class Poitest {

    private static String CONTROL_SHEET_NAME = "Kontrolle";
    private static String REC_SHEET_NAME = "Anweisungen";
    private static String DIV_SHEET_NAME = "Divers";

    private String controlFile;
    
//    private List<Posten> posten;

    public Poitest() throws IOException, InvalidFormatException {
        controlFile = copyWorkBook("C:\\Users\\rma\\Desktop\\BBH\\Baukontrolle_AMSEL_sevi_roman_neu.xlsx");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        new Poitest().run();
    }

    private void run() throws FileNotFoundException, IOException, InvalidFormatException {

        Workbook wb = new XSSFWorkbook(new File(controlFile));
//        Sheet s = wb.getSheet(CONTROL_SHEET_NAME);
//        Row r = s.getRow(387);
//        Cell c = r.getCell(15);
//        System.out.println(c.getNumericCellValue());
//        System.out.println(c.getCellStyle().getDataFormatString());
//        System.out.println(c.getCellStyle().getDataFormat());

        List<Posten> posten = new ArrayList<>();

        posten.addAll(getAnweisungen(wb, REC_SHEET_NAME));
        posten.addAll(getAnweisungen(wb, DIV_SHEET_NAME));

        wb.close();

        for (Posten p : posten) {
            System.out.println(p);
            addToSheet(p);
        }
    }

    private List<Posten> getAnweisungen(Workbook wb, String sheetName) {
        List<Posten> declined = new ArrayList<>();
        List<Posten> rowInControl = new ArrayList<>();

        Sheet s = wb.getSheet(sheetName);
        Iterator<Row> rows = s.rowIterator();

        rows.next();
        while (rows.hasNext()) {
            Row r = rows.next();
            Posten p = parseRekapLine(r);
            if (p.getBkp().equals("Divers") || p.getKostenstelle().equals("Divers")) {
                declined.add(p);
            } else {
                rowInControl.add(p);
            }
        }

        return rowInControl;
    }

    private void addToSheet(Posten p) throws IOException, InvalidFormatException {
        Workbook wb = new XSSFWorkbook(new FileInputStream(controlFile));
        Sheet s = wb.getSheet(CONTROL_SHEET_NAME);
        List<Kostenstelle> kostenstellen = loadKsTitles(s);
        kostenstellen.addAll(loadKs(s));
        Kostenstelle KostenstelleVonP = findKs(p, kostenstellen);
        s.shiftRows(KostenstelleVonP.getFirstLineInSheet() + 1, s.getLastRowNum(), 1);
        Row newRow = s.createRow(KostenstelleVonP.getFirstLineInSheet() + 1);

        CellStyle cellStyle = wb.createCellStyle();
        CreationHelper createHelper = wb.getCreationHelper();
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("\"CHF\"\\ #,##0.00"));

        Cell summe = newRow.createCell(13);
        summe.setCellStyle(cellStyle);
        summe.setCellValue(p.getSumme());
        setCellFontTo8Pt(summe);

        Cell bkp = newRow.createCell(0);
        bkp.setCellValue(p.getBkp());
        setCellFontTo8Pt(bkp);

        Cell ks = newRow.createCell(1);
        ks.setCellValue(p.getKostenstelle());
        setCellFontTo8Pt(ks);

        Cell empfaenger = newRow.createCell(7);
        empfaenger.setCellValue(p.getEmpfaenger());
        setCellFontTo8Pt(empfaenger);

        Cell reNr = newRow.createCell(12);
        reNr.setCellValue(p.getReNr());
        setCellFontTo8Pt(reNr);

        wb.write(new FileOutputStream(controlFile));
        wb.setForceFormulaRecalculation(true);
        wb.close();
    }

    private Kostenstelle findKs(Posten p, List<Kostenstelle> kostenstellen) {
        Optional<Kostenstelle> ks = kostenstellen.stream().filter(k -> {
            return p.getBkp().equals(k.getBkpNr().toString())
                    && p.getKostenstelle().equals(k.getKs());
        })
                .findFirst();

        if (!ks.isPresent()) {
            System.out.println("Kostenstelle nicht gefunden: " + p);
        }

        return ks.get();
    }

    private String copyWorkBook(String file) throws IOException, InvalidFormatException {
        int fileEndingPos = file.toLowerCase().lastIndexOf(".xlsx");
        String timeStamp = new SimpleDateFormat("ddMMyyyyHHmmss").format(new java.util.Date());
        String newFileName = file.substring(0, fileEndingPos) + "_" + timeStamp + file.substring(fileEndingPos);

        Files.copy(Paths.get(file), Paths.get(newFileName));

//        Workbook wb = new XSSFWorkbook(new File(file));
//        Sheet s = wb.getSheet(CONTROL_SHEET_NAME);
//        fillKs(s);
//        wb.write(new FileOutputStream(newFileName));
        return newFileName;
    }

    private void fillKs(Sheet s) {
        Kostenstelle kostenstelle = null;

        Iterator<Row> rows = s.rowIterator();
        while (rows.hasNext()) {
            Row r = rows.next();
            Integer bkpNr = null;
            String ksNo;
            String name = "";

            if (r.getCell(0) == null || r.getCell(1) == null) {
                continue;
            }

            if (r.getCell(0).getCellType().equals(CellType.NUMERIC)) {
                bkpNr = (int) r.getCell(0).getNumericCellValue();
            }

            if (r.getCell(2) != null) {
                name = r.getCell(2).toString();
            }

            Cell ksCell = r.getCell(1);

            DataFormatter formatter = new DataFormatter();
            ksNo = formatter.formatCellValue(ksCell);

            if (ksNo != null && !ksNo.equals("") && bkpNr != null && bkpNr != 0) {
                Kostenstelle ks = new Kostenstelle(bkpNr, ksNo, name, r.getRowNum());
                kostenstelle = ks;
            } else if (kostenstelle != null && bkpNr == null && (ksNo == null || ksNo.equals(""))) {
                Cell bkp = r.createCell(0);
                bkp.setCellValue(kostenstelle.getBkpNr());
                setCellFontTo8Pt(bkp);

                Cell ks = r.createCell(1);
                ks.setCellValue(kostenstelle.getKs());
                setCellFontTo8Pt(ks);
            }
        }
    }

    private void setCellFontTo8Pt(Cell c) {
        Font newFont = c.getSheet().getWorkbook().createFont();
        newFont.setFontHeightInPoints((short) 8);
        CellUtil.setFont(c, newFont);
    }

    private Posten parseRekapLine(Row row) {
        DataFormatter formatter = new DataFormatter();

        Double summe = row.getCell(5).getNumericCellValue();

        Cell bkpCell = row.getCell(0);
        String bkp = formatter.formatCellValue(bkpCell);

        Cell reCell = row.getCell(2);
        String reNr = formatter.formatCellValue(reCell);

        Cell ksCell = row.getCell(1);
        String ks = formatter.formatCellValue(ksCell);

        Cell empfaengerCell = row.getCell(4);
        String empfaenger = formatter.formatCellValue(empfaengerCell);

        Cell kommentarCell = row.getCell(7);
        String kommentar = formatter.formatCellValue(kommentarCell);

        return new Posten(summe, bkp, ks, empfaenger, reNr, 0, kommentar);
    }

    private List<Kostenstelle> loadKs(Sheet s) {
        List<Kostenstelle> kostenstellen = new ArrayList<>();

        Iterator<Row> rows = s.rowIterator();
        while (rows.hasNext()) {
            Row r = rows.next();
            Integer bkpNr = 0;
            String ksNo = "";
            String name = "";

            if (r.getCell(0) == null || r.getCell(1) == null) {
                continue;
            }

            if (r.getCell(0).getCellType().equals(CellType.NUMERIC)) {
                bkpNr = (int) r.getCell(0).getNumericCellValue();
            }

            if (r.getCell(2) != null) {
                name = r.getCell(2).toString();
            }

            Cell ksCell = r.getCell(1);

            DataFormatter formatter = new DataFormatter();
            ksNo = formatter.formatCellValue(ksCell);

            if (ksNo != null && !ksNo.equals("") && bkpNr != 0 && ksNo.matches("[^a-zA-Z]+")) {
                if (kostenstellen.size() > 0) {
                    kostenstellen.get(kostenstellen.size() - 1).setLastLineInSheet(r.getRowNum() - 1);
                }
                Kostenstelle ks = new Kostenstelle(bkpNr, ksNo, name, r.getRowNum());
                kostenstellen.add(ks);
            }
        }
        return kostenstellen;
    }

    private List<Kostenstelle> loadKsTitles(Sheet s) {
        List<Kostenstelle> kostenstellen = new ArrayList<>();

        Iterator<Row> rows = s.rowIterator();
        while (rows.hasNext()) {
            Row r = rows.next();
            Integer bkpNr = 0;
            String ksNo = "";
            String name = "";

            if (r.getCell(0) == null || r.getCell(1) == null) {
                continue;
            }

            if (r.getCell(0).getCellType().equals(CellType.NUMERIC)) {
                bkpNr = (int) r.getCell(0).getNumericCellValue();
            }

            if (r.getCell(2) != null) {
                name = r.getCell(2).toString();
            }

            Cell ksCell = r.getCell(1);

            DataFormatter formatter = new DataFormatter();
            ksNo = formatter.formatCellValue(ksCell);

            if (ksNo != null && !ksNo.equals("") && bkpNr != 0 && ksNo.matches("[^0-9]+")) {
                if (kostenstellen.size() > 0) {
                    kostenstellen.get(kostenstellen.size() - 1).setLastLineInSheet(r.getRowNum() - 1);
                }
                Kostenstelle ks = new Kostenstelle(bkpNr, "", ksNo, r.getRowNum());
                kostenstellen.add(ks);
            }
        }
        return kostenstellen;
    }
}
