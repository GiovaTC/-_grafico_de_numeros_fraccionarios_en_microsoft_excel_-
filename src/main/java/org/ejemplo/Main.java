package org.ejemplo;

import java.util.Scanner;

public class Main {

    public static void main(String[] args) {

        Scanner sc = new Scanner(System.in);
        double[] numeros = new double[4];

        System.out.println("INGRESE 4 NUMEROS FRACCIONARIOS");

        for (int i = 0; i < numeros.length; i++) {

            System.out.print("NUMERO " + (i + 1) + ": ");
            numeros[i] = sc.nextDouble();

        }
        ExcelService.generarExcel(numeros);

        System.out.println();
        System.out.println("archivo excel generado correctamente! ");
    }
}