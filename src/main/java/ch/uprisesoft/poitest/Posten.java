/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ch.uprisesoft.poitest;

/**
 *
 * @author rma
 */
public class Posten {
    private final Double summe;
    private final String bkp;
    private final String kostenstelle;
    private final String reNr;
    private final String empfaenger;
    private final int rekap;
    private final String kommentar;

    public Posten(Double summe, String bkp, String kostenstelle,  String empfaenger, String reNr, int rekap, String kommentar) {
        this.summe = summe;
        this.bkp = bkp;
        this.kostenstelle = kostenstelle;
        this.empfaenger = empfaenger;
        this.reNr = reNr;
        this.rekap = rekap;
        this.kommentar = kommentar;
    }

    public Double getSumme() {
        return summe;
    }

    public String getBkp() {
        return bkp;
    }

    public String getKostenstelle() {
        return kostenstelle;
    }

    public String getEmpfaenger() {
        return empfaenger;
    }

    public String getReNr() {
        return reNr;
    }

    public int getRekap() {
        return rekap;
    }

    public String getKommentar() {
        return kommentar;
    }

    @Override
    public String toString() {
        return "Posten{" + "summe=" + summe + ", bkp=" + bkp + ", kostenstelle=" + kostenstelle + ", reNr=" + reNr + ", empfaenger=" + empfaenger + ", rekap=" + rekap + ", kommentar=" + kommentar + '}';
    }
    
}
