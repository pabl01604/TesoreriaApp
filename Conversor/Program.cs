using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;

namespace Conversor
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Rutas de los archivos de entrada y salida
            string rutaDiaAnterior = @"C:\Users\Pablo\Documents\GGARCIA\Conversion de excels\extractos.xlsx";
            string rutaDiaActual = @"C:\Users\Pablo\Documents\GGARCIA\Conversion de excels\movimientos.xlsx";
            string rutaArchivoSalida = @"C:\Users\Pablo\Documents\GGARCIA\Conversion de excels\MovimientosCombinados.xlsx";


            // Leer los datos de ambos archivos
            List<Movimiento> movimientosDiaAnterior = LeerMovimientosDesdeExcel(rutaDiaAnterior, "anterior");
            List<Movimiento> movimientosDiaActual = LeerMovimientosDesdeExcel(rutaDiaActual, "hoy");

            // Combinar los movimientos
            List<Movimiento> movimientosCombinados = new List<Movimiento>();
            movimientosCombinados.AddRange(movimientosDiaAnterior);
            movimientosCombinados.AddRange(movimientosDiaActual);

            // Generar el archivo Excel combinado
            GenerarExcelConMovimientos(movimientosCombinados, rutaArchivoSalida);

            Console.WriteLine($"Archivo combinado generado en: {rutaArchivoSalida}");
            Console.ReadKey();
        }

        private static void GenerarExcelConMovimientos(List<Movimiento> movimientos, string rutaArchivoSalida)
        {
            FileInfo archivo = new FileInfo(rutaArchivoSalida);

            using (ExcelPackage paquete = new ExcelPackage())
            {
                ExcelWorksheet hoja = paquete.Workbook.Worksheets.Add("Movimientos Combinados");

                // Encabezados
                hoja.Cells[1, 1].Value = "Fecha";
                hoja.Cells[1, 2].Value = "Suc Origen";
                hoja.Cells[1, 3].Value = "Desc Sucursal";
                hoja.Cells[1, 4].Value = "Cod Operativo";
                hoja.Cells[1, 5].Value = "Referencia";
                hoja.Cells[1, 6].Value = "Concepto";
                hoja.Cells[1, 7].Value = "Importe Pesos";

                //movimientos.RemoveAll(movimiento => movimiento.Concepto == "DEPOSITO E-CHEQ INT MISMA PLAZA");
                movimientos.RemoveAll(movimiento => movimiento.ImporteEnPesos < 0);

                //Ordenar por fechas
                var movimientosOrdenados = movimientos.OrderBy(m => m.Fecha).ToList();
                // Escribir los movimientos
                int fila = 2;
                foreach (var movimiento in movimientosOrdenados)
                {
                    if (movimiento.ImporteEnPesos > 0 && !movimiento.Concepto.Trim().ToUpper().Contains("E-CHEQ"))
                    {
                        hoja.Cells[fila, 1].Value = movimiento.Fecha.ToString("dd/MM/yyyy");
                        hoja.Cells[fila, 2].Value = movimiento.SucOrigen;
                        hoja.Cells[fila, 3].Value = movimiento.DescSucursal;
                        hoja.Cells[fila, 4].Value = movimiento.CodOperativo;
                        hoja.Cells[fila, 5].Value = movimiento.Referencia;
                        hoja.Cells[fila, 6].Value = movimiento.Concepto;
                        hoja.Cells[fila, 7].Value = movimiento.ImporteEnPesos;
                    }
                    fila++;
                }

                // Ajustar ancho de columnas automáticamente
                hoja.Cells[hoja.Dimension.Address].AutoFitColumns();

                // Guardar el archivo
                paquete.SaveAs(archivo);
            }
        }

        private static List<Movimiento> LeerMovimientosDesdeExcel(string rutaArchivo, string dia)
        {
            List<Movimiento> movimientos = new List<Movimiento>();

            FileInfo archivo = new FileInfo(rutaArchivo);
            if (dia == "hoy")
            {
                using (ExcelPackage paquete = new ExcelPackage(archivo))
                {
                    ExcelWorksheet hoja = paquete.Workbook.Worksheets[0]; // Leer la primera hoja
                    int filas = hoja.Dimension.Rows;

                    // Asumimos que la primera fila es el encabezado
                    for (int fila = 2; fila <= filas; fila++) // Comenzar desde la fila 2
                    {
                        /*
                        DateTime fecha = DateTime.Parse(hoja.Cells[fila, 1].Text); // Columna 1: Fecha
                        string numeroTransaccion = hoja.Cells[fila, 2].Text;       // Columna 2: Número de Transacción
                        decimal monto = decimal.Parse(hoja.Cells[fila, 3].Text);  // Columna 3: Monto
                        string descripcion = hoja.Cells[fila, 4].Text;           // Columna 4: Descripción
                        */

                        string fechaEntrada = hoja.Cells[fila, 1].Text; // Columna 1: Fecha
                        DateTime fecha = DateTime.Parse(fechaEntrada);

                        int sucursal = Convert.ToInt32(hoja.Cells[fila, 2].Text); // Columna 2: Sucursal

                        int codOperativo = Convert.ToInt32(hoja.Cells[fila, 4].Text);  // Columna 4: Cod operativo

                        int referencia = Convert.ToInt32(hoja.Cells[fila, 3].Text);// Columna 3: Referencia

                        string concepto = hoja.Cells[fila, 5].Text; // Columna 5: concepto

                        double importe = Convert.ToDouble(hoja.Cells[fila, 6].Text); // Columna 6: importe

                        movimientos.Add(new Movimiento(fecha, sucursal, codOperativo, referencia, concepto, importe));


                    }
                }

            }
            else if (dia == "anterior")
            {
                using (ExcelPackage paqueteDos = new ExcelPackage(archivo))
                {
                    ExcelWorksheet hoja = paqueteDos.Workbook.Worksheets[0]; // Leer la primera hoja
                    int filas = hoja.Dimension.Rows;

                    // Asumimos que la primera fila es el encabezado
                    for (int fila = 2; fila <= filas; fila++) // Comenzar desde la fila 2
                    {
                        /*
                        DateTime fecha = DateTime.Parse(hoja.Cells[fila, 1].Text); // Columna 1: Fecha
                        string numeroTransaccion = hoja.Cells[fila, 2].Text;       // Columna 2: Número de Transacción
                        decimal monto = decimal.Parse(hoja.Cells[fila, 3].Text);  // Columna 3: Monto
                        string descripcion = hoja.Cells[fila, 4].Text;           // Columna 4: Descripción
                        */

                        string fechaEntrada = hoja.Cells[fila, 1].Text; // Columna 1: Fecha
                        //DateTime fecha = DateTime.Parse(fechaEntrada);
                        DateTime fecha = DateTime.ParseExact(fechaEntrada, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        int sucursal = Convert.ToInt32(hoja.Cells[fila, 5].Text); // Columna 2: Sucursal
                 


                        int codOperativo = Convert.ToInt32(hoja.Cells[fila, 7].Text);  // Columna 4: Cod operativo

                        int referencia = Convert.ToInt32(hoja.Cells[fila, 4].Text);// Columna 3: Referencia

                        string concepto = hoja.Cells[fila,2].Text; // Columna 5: concepto

                        double importe = Convert.ToDouble(hoja.Cells[fila, 3].Text); // Columna 6: importe

                        movimientos.Add(new Movimiento(fecha, sucursal, codOperativo, referencia, concepto, importe));

                    }
                }       
            }
            return movimientos;
        }
    }
}
