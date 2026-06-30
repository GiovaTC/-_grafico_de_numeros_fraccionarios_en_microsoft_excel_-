# -_grafico_de_numeros_fraccionarios_en_microsoft_excel_- :.
# Proyecto Java 21 + IntelliJ IDEA + Apache POI:

<img width="1254" height="1254" alt="image" src="https://github.com/user-attachments/assets/0245e552-8b1d-4697-9d45-badbea0b0366" />         

<img width="2559" height="1032" alt="image" src="https://github.com/user-attachments/assets/76e331a6-deb6-4ef1-a543-717ac828077c" />         

```
# Gráfico de Números Fraccionarios en Microsoft Excel (.xlsx)
Proyecto desarrollado en **Java 21** utilizando **Apache POI**, que permite almacenar números fraccionarios en un
archivo de **Microsoft Excel (.xlsx)** y generar automáticamente un **gráfico de barras** dentro de la misma hoja de calculo .

---

# Caracteristicas:

* ✅ Solicita **4 números fraccionarios**.
* ✅ Guarda los datos en un archivo **Excel (.xlsx)**.
* ✅ Genera automáticamente un **gráfico de barras**.
* ✅ Compatible con **Java 21**.
* ✅ Desarrollado en **IntelliJ IDEA**.
* ✅ Utiliza la librería **Apache POI**.

---

# Tecnologías utilizadas

* Java 21
* Apache Maven
* Apache POI 5.4.1
* Microsoft Excel (.xlsx)
* IntelliJ IDEA

---

# Estructura del Proyecto

```text
GraficoFraccionariosExcel
│
├── src
│   └── main
│       └── java
│            └── org
│                 └── ejemplo
│                      ├── Main.java
│                      └── ExcelService.java
│
├── pom.xml
│
└── resultados.xlsx
```

---

# Archivo `pom.xml`

```xml
<project xmlns="http://maven.apache.org/POM/4.0.0"
         xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
         xsi:schemaLocation="
         http://maven.apache.org/POM/4.0.0
         https://maven.apache.org/xsd/maven-4.0.0.xsd">

    <modelVersion>4.0.0</modelVersion>

    <groupId>org.ejemplo</groupId>
    <artifactId>GraficoFraccionariosExcel</artifactId>
    <version>1.0</version>

    <properties>
        <maven.compiler.source>21</maven.compiler.source>
        <maven.compiler.target>21</maven.compiler.target>
    </properties>

    <dependencies>

        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>5.4.1</version>
        </dependency>

    </dependencies>

</project>
```

---

# Clase `Main.java`

```java
package org.ejemplo;

import java.util.Scanner;

public class Main {

    public static void main(String[] args) {

        Scanner sc = new Scanner(System.in);

        double[] numeros = new double[4];

        System.out.println("INGRESE 4 NUMEROS FRACCIONARIOS");

        for (int i = 0; i < numeros.length; i++) {

            System.out.print("Numero " + (i + 1) + ": ");
            numeros[i] = sc.nextDouble();

        }

        ExcelService.generarExcel(numeros);

        System.out.println();
        System.out.println("Archivo Excel generado correctamente.");

    }

}
```

---

# Clase `ExcelService.java`

```java
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

            for (int i = 0; i < numeros.length; i++) {

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
```

---

# Funcionamiento del Programa

El programa realiza las siguientes operaciones:

1. Solicita **4 números fraccionarios**.
2. Crea un archivo **Excel (.xlsx)**.
3. Inserta los datos en una tabla.
4. Genera automáticamente un **gráfico de barras**.
5. Guarda el archivo como:

```text
resultados.xlsx
```

---

# Resultado en Microsoft Excel

La hoja generada contendrá la siguiente información:

| Número | Valor |
| ------ | ----: |
| Dato 1 |  2.35 |
| Dato 2 |  8.40 |
| Dato 3 |  5.70 |
| Dato 4 |  9.15 |

---

# Gráfico generado

```text
Grafico de Numeros Fraccionarios

10 ┤                           █
 9 ┤                           █
 8 ┤              █            █
 7 ┤              █            █
 6 ┤              █            █
 5 ┤      █       █      █     █
 4 ┤      █       █      █     █
 3 ┤      █       █      █     █
 2 ┤ █    █       █      █     █
 1 ┤ █    █       █      █     █
 0 └────────────────────────────────
     D1   D2      D3     D4
```

---

# Ejemplo de ejecución

```text
INGRESE 4 NUMEROS FRACCIONARIOS

Numero 1: 2.75
Numero 2: 5.60
Numero 3: 3.20
Numero 4: 8.95

Archivo Excel generado correctamente.
```

---

# Archivo generado

```text
resultados.xlsx
```

Contenido del archivo:

* Tabla con los cuatro números fraccionarios.
* Gráfico de barras incrustado.
* Datos compatibles con Microsoft Excel.
* Archivo listo para abrir sin configuraciones adicionales.

---

# Requisitos

* Java 21
* Apache Maven
* IntelliJ IDEA
* Apache POI 5.4.1
* Microsoft Excel (opcional para visualizar el archivo)

---

# Compilar el proyecto

```bash
mvn clean compile
```

---

# Ejecutar el proyecto

```bash
mvn exec:java
```

o ejecutar directamente la clase:

```text
Main.java
```

desde IntelliJ IDEA.

---

# Resultado esperado

Al finalizar la ejecución se obtendrá:

* ✔ Archivo `resultados.xlsx`
* ✔ Tabla con los datos ingresados
* ✔ Gráfico de barras generado automáticamente
* ✔ Compatible con Microsoft Excel
* ✔ Compatible con Java 21
* ✔ Código organizado mediante clases independientes

---

# Posibles mejoras

El proyecto puede ampliarse fácilmente para incluir:

* 📈 Gráfico de líneas.
* 📊 Gráfico de columnas.
* 🥧 Gráfico circular.
* 📉 Gráfico de dispersión.
* 📋 Múltiples hojas de cálculo.
* 📅 Registro de fecha y hora.
* 🎨 Aplicación de estilos y formatos personalizados.
* 📄 Exportación de múltiples reportes en Excel.

---

# Autor

Proyecto desarrollado en **Java 21** utilizando **Apache POI** para la generación automática de archivos de Microsoft Excel con gráficos integrados .
:. . / .
