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
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import javax.swing.plaf.basic.BasicBorders;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
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

    public Poitest() throws IOException, InvalidFormatException {
        controlFile = copyWorkBook("C:\\Users\\rma\\Desktop\\BBH\\Baukontrolle_AMSEL_sevi_roman_test.xlsx");
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        new Poitest().run();
    }

    private void run() throws FileNotFoundException, IOException, InvalidFormatException {

        Workbook wb = new XSSFWorkbook(new File(controlFile));

        List<Posten> posten = new ArrayList<>();
        
        posten.addAll(getAnweisungen(wb, REC_SHEET_NAME));
        posten.addAll(getAnweisungen(wb, DIV_SHEET_NAME));
        
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
        
        newRow.createCell(13).setCellValue(p.getSumme()); // Summe
        newRow.createCell(0).setCellValue(p.getBkp()); // BKP
        newRow.createCell(1).setCellValue(p.getKostenstelle()); // Kostenstelle
        newRow.createCell(7).setCellValue(p.getEmpfaenger()); // Empfänger
        newRow.createCell(12).setCellValue(p.getReNr()); // Beilage-Nr.
        wb.write(new FileOutputStream(controlFile));
    }

    private Kostenstelle findKs(Posten p, List<Kostenstelle> kostenstellen) {
        return kostenstellen.stream().filter(k -> {
            return p.getBkp().equals(k.getBkpNr().toString())
                    && p.getKostenstelle().equals(k.getKs());
        })
                .findFirst().get();
    }

    private String copyWorkBook(String file) throws IOException, InvalidFormatException {
        int fileEndingPos = file.toLowerCase().lastIndexOf(".xlsx");
        String timeStamp = new SimpleDateFormat("ddMMyyyyHHmmss").format(new java.util.Date());
        String newFileName = file.substring(0, fileEndingPos) + "_" + timeStamp + file.substring(fileEndingPos);

        Workbook wb = new XSSFWorkbook(new File(file));
        wb.write(new FileOutputStream(newFileName));

        return newFileName;
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
        
//        int rekap = (int) row.getCell(6).getNumericCellValue();
        
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
