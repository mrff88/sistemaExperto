using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SistemasExpertos
{
    class Reporte
    {
        public Reporte()
        {

        }

        public static void Generar(List<Fiscal> Fiscalia)
        {
            SLDocument reporte = new SLDocument();//creamos una instancia del documento excel para manipularlo
            int numColDoc = (Fiscalia.Count*4)+7;//numero de colummnas necesarias
            //cabezera
            SLStyle cabezera = reporte.CreateStyle();//creamos un estilo
            cabezera.SetWrapText(true);//activamos el ajuste de texto
            cabezera.Font.FontName = "Arial";
            cabezera.Font.Bold = true;//activamos que las letras esten en negrita
            cabezera.SetHorizontalAlignment(HorizontalAlignmentValues.Center);//centramos el contenido de la celda
            cabezera.SetVerticalAlignment(VerticalAlignmentValues.Center);//centramos el contenido de la celda
            reporte.SetCellStyle("A1", cabezera);//damos el estilo a la celda A1
            reporte.SetCellValue("A1", "DISTRITO JUDICIAL DE MADRE DE DIOS\nEVOLUCION PERIODO\nFISCALIA ESPECIALIZADA EN CORRUPCION DE FUNCIONARIOS");
            reporte.MergeWorksheetCells(1, 1, 1, numColDoc);//mezclamos las celdas
            reporte.SetRowHeight(1,45);//damos un alto de 45 a la fila 1

            reporte.SetCellStyle("A2", cabezera);//damos el estilo a la celda A2
            reporte.SetCellValue("A2", "PERIODO "+" A ");
            reporte.MergeWorksheetCells(2, 1, 2, numColDoc);//mezclamos las celdas
            //cuerpo
            int contador = 4;
            foreach (Fiscal fisc in Fiscalia)
            {
                reporte.SetCellStyle(4, contador, cabezera);//damos el estilo a la celda
                reporte.SetCellValue(4, contador, fisc.Apell + ", " + fisc.Nomb);
                reporte.MergeWorksheetCells(4, contador, 4, (contador+3));//mezclamos las celdas
                reporte.SetRowHeight(4, 45);//damos un alto de 45 a la fila 4
                contador += 4;
            }
            reporte.SaveAs("prueba01.xlsx");//guardamos el documento
        }
    }
}
