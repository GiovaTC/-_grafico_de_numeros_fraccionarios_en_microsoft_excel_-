package org.ejemplo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;

public class ExcelService {

    public static void generarExcel(double[] numeros) {

        try {
            XSSFWorkbook libro = new XSSFWorkbook();
            XSSFSheet hoja = libro.createSheet("Datos");
            Row encabezado = hoja.createRow(0);

            encabezado.createCell(0).setCellValue("Numero");
            encabezado.createCell(1).setCellValue("Valor");

            for (int i=0; i < numeros.length; i++) {
                Row fila = hoja.createRow(i + 1);

                fila.createCell(0).setCellValue("Dato " + (i + 1));
                fila.createCell(1).setCellValue(numeros[i]);
            }

            XSSFDrawing dibujo = hoja.createDrawingPatriarch();
            XSSFClientAnchor ancla = dibujo.createAnchor(
                    0,0,0,0,
                    3,1,
                    12,20);
            XSSFChart grafico = dibujo.createChart(ancla);

            grafico.setTitleText("Grafico de Numeros Fraccionarios");
            grafico.setTitleOverlay(false);

            XDDFCategoryAxis ejeX =
                    grafico.createCategoryAxis(AxisPosition.BOTTOM);

            ejeX.setTitle("Datos");

            XDDFValueAxis ejeY =
                    grafico.createValueAxis(AxisPosition.LEFT);

            ejeY.setTitle("Valor");

            XDDFDataSource<String> categorias =
                    XDDFDataSourcesFactory.fromStringCellRange(
                            hoja,
                            new CellRangeAddress(1,4,0,0));

            XDDFNumericalDataSource<Double> valores =
                    XDDFDataSourcesFactory.fromNumericCellRange(
                            hoja,
                            new CellRangeAddress(1,4,1,1));

            XDDFBarChartData datos =
                    (XDDFBarChartData) grafico.createData(
                            ChartTypes.BAR,
                            ejeX,
                            ejeY);

            XDDFBarChartData.Series serie =
                    (XDDFBarChartData.Series)
                            datos.addSeries(categorias, valores);

            serie.setTitle("Valores", null);

            grafico.plot(datos);

            FileOutputStream archivo =
                    new FileOutputStream("resultados.xlsx");

            libro.write(archivo);

            archivo.close();
            libro.close();

        } catch (Exception e) {

            e.printStackTrace();
        }
    }   
}
