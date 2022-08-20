/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Project/Maven2/JavaApp/src/main/java/${packagePath}/${mainClassName}.java to edit this template
 */
package ch.uprisesoft.poitest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
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
        controlFile = copyWorkBook("C:\\Users\\roman\\OneDrive\\Desktop\\BBH\\Baukontrolle_AMSEL_sevi_roman_neu_08082022154824_korrigiert.xlsx");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException, FileNotFoundException, CloneNotSupportedException {
        new Poitest().run();
    }

    private void run() throws FileNotFoundException, IOException, InvalidFormatException, CloneNotSupportedException {

        Workbook wb = new XSSFWorkbook(new FileInputStream(controlFile));
//        Sheet s = wb.getSheet(CONTROL_SHEET_NAME);
//        Row r = s.getRow(387);
//        Cell c = r.getCell(15);
//        System.out.println(c.getNumericCellValue());
//        System.out.println(c.getCellStyle().getDataFormatString());
//        System.out.println(c.getCellStyle().getDataFormat());

        List<Posten> posten = new ArrayList<>();

        posten.addAll(getAnweisungen(wb, REC_SHEET_NAME));
        posten.addAll(getAnweisungen(wb, DIV_SHEET_NAME));

        List<Posten> postenGruppen = new ArrayList<>();

        for (Posten p : posten) {
            groupPosten(postenGruppen, p);
        }

//        wb.close();
//        wb = new XSSFWorkbook(new FileInputStream(controlFile));
//        ZipSecureFile.setMinInflateRatio(-1.0d);
//        for (Posten p : posten) {
//            System.out.println(p);
//            addToSheet(p, wb);
//        }
        for (Posten p : postenGruppen) {
            System.out.println(p);
            addToSheet(p, wb);
        }

//        BigDecimal sum = new BigDecimal(0);
//        for (Posten p : posten) {
//            sum = sum.add(p.getSumme());
//        }
//        System.out.println(sum);
//        
//        sum = new BigDecimal(0);
//        for (Posten p : postenGruppen) {
//            sum = sum.add(p.getSumme());
//        }
//        System.out.println(sum);
//        
//        
//        System.out.println(posten.size());
//        System.out.println(postenGruppen.size());
        // TODO
        wb.write(new FileOutputStream(controlFile));
        wb.setForceFormulaRecalculation(true);
        wb.close();
    }

    private void groupPosten(List<Posten> postenSummen, Posten posten) throws CloneNotSupportedException {

        for (Posten gp : postenSummen) {
            if (gp.getBkp().equals(posten.getBkp())
                    && gp.getKostenstelle().equals(posten.getKostenstelle())
                    && gp.getReNr().equals(posten.getReNr())) {
//                System.out.println("Match:");
//                System.out.println(posten);
//                System.out.println(gp);
//                System.out.println("===================");
                gp.setSumme(gp.getSumme().add(posten.getSumme()));
                return;
            }
        }
        postenSummen.add(posten.clone());
    }

    private List<Posten> getAnweisungen(Workbook wb, String sheetName) {
//        List<Posten> declined = new ArrayList<>();
        List<Posten> rowInControl = new ArrayList<>();

        Sheet s = wb.getSheet(sheetName);
        Iterator<Row> rows = s.rowIterator();

        rows.next();
        while (rows.hasNext()) {
            Row r = rows.next();
            try {
                Posten p = parseRekapLine(r);
                if (p.getBkp().toLowerCase().equals("divers") || p.getKostenstelle().toLowerCase().equals("divers")) {
//                declined.add(p);
                } else {
                    rowInControl.add(p);
                }
            } catch (NullPointerException npe) {
                System.out.println("NPE in line " + r.getRowNum() + ", Sheet " + r.getSheet().getSheetName());
                System.exit(-1);
            }

        }

        return rowInControl;
    }

    private void addToSheet(Posten p, Workbook wb) throws IOException, InvalidFormatException {
//        Workbook wb = new XSSFWorkbook(new FileInputStream(controlFile));
//        ZipSecureFile.setMinInflateRatio(-1.0d); 
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
        summe.setCellValue(p.getSumme().doubleValue());
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

//        Cell kommentar = newRow.createCell(8);
//        kommentar.setCellValue(p.getKommentar());
//        setCellFontTo8Pt(kommentar);
        Cell rekap = newRow.createCell(20);
        rekap.setCellValue(p.getRekap());
        setCellFontTo8Pt(rekap);

//        wb.write(new FileOutputStream(controlFile));
//        wb.setForceFormulaRecalculation(true);
//        wb.close();
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

    private Posten parseRekapLine(Row row) throws NullPointerException {
        DataFormatter formatter = new DataFormatter();

        String summe = Double.toString(row.getCell(5).getNumericCellValue());

        Cell bkpCell = row.getCell(0);
        String bkp = formatter.formatCellValue(bkpCell);

        Cell reCell = row.getCell(2);
        String reNr = formatter.formatCellValue(reCell);

        Cell ksCell = row.getCell(1);
        String ks = formatter.formatCellValue(ksCell);

        Cell empfaengerCell = row.getCell(3);
        String empfaenger = formatter.formatCellValue(empfaengerCell);

//        Cell kommentarCell = row.getCell(7);
//        String kommentar = formatter.formatCellValue(kommentarCell);
        Cell rekapCell = row.getCell(6);
        String rekap = formatter.formatCellValue(rekapCell);

        return new Posten(summe, bkp, ks, empfaenger, reNr, rekap);
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
