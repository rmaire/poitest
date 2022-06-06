/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package ch.uprisesoft.poitest;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 *
 * @author rma
 */
public class Kostenstelle {
    
    private final Integer bkpNr;
    private final String ks;
    private final String name;
    private final int firstLineInSheet;
    private int lastLineInSheet;
//    private final List<Integer> ranks = new ArrayList<>();
    
    public Kostenstelle(Integer bkpNr, String ks, String name, int firstLine) {
        this.bkpNr = bkpNr;
        this.ks = ks.trim();
        this.name = name.trim();
        this.firstLineInSheet = firstLine;
        this.lastLineInSheet = firstLine;
    }

    public void setLastLineInSheet(int lastLineInSheet) {
        this.lastLineInSheet = lastLineInSheet;
    }

    public int getFirstLineInSheet() {
        return firstLineInSheet;
    }

    public int getLastLineInSheet() {
        return lastLineInSheet;
    }
    
    

    public Integer getBkpNr() {
        return bkpNr;
    }

    public String getKs() {
        return ks;
    }
//
//    public List<Integer> getRanks() {
//        return ranks;
//    }

    public String getName() {
        return name;
    }
    
    public String fullStr(){
        StringBuilder full = new StringBuilder();
        return full.append(bkpNr).append(" ")
                .append(ks).append(" ")
                .append(name)
                .toString();
    }
    
    public String fullStrCsv(){
        StringBuilder full = new StringBuilder();
        return full.append(bkpNr).append(";")
                .append(ks).append(";")
                .append(name).append(";")
                .append(firstLineInSheet).append(";")
                .append(lastLineInSheet)
                .toString();
    }

    @Override
    public String toString() {
        return "Kostenstelle{" + "bkpNr=" + bkpNr + ", ks=" + ks + ", name=" + name + ", firstLineInSheet=" + firstLineInSheet + ", lastLineInSheet=" + lastLineInSheet + '}';
    }
}
