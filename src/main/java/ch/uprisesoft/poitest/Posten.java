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

    public Posten(Double summe, String bkp, String kostenstelle,  String empfaenger, String reNr, int rekap) {
        this.summe = summe;
        this.bkp = bkp;
        this.kostenstelle = kostenstelle;
        this.empfaenger = empfaenger;
        this.reNr = reNr;
        this.rekap = rekap;
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

    @Override
    public String toString() {
        return "Posten{" + "summe=" + summe + ", bkp=" + bkp + ", empfaenger=" + empfaenger + ", reNr=" + reNr + ", rekap=" + rekap + '}';
    }

    
}
