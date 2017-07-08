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
            int numColDoc = (Fiscalia.Count * 4) + 7;//numero de colummnas necesarias
            Fiscal totalFiscal= new Fiscal();
            //ESTILOS
            //--estilo cabezera
            SLStyle cabezera = reporte.CreateStyle();//creamos un estilo
            cabezera.SetWrapText(true);//activamos el ajuste de texto
            cabezera.Font.FontName = "Arial";
            cabezera.Font.Bold = true;//activamos que las letras esten en negrita
            cabezera.SetHorizontalAlignment(HorizontalAlignmentValues.Center);//centramos el contenido de la celda
            cabezera.SetVerticalAlignment(VerticalAlignmentValues.Center);//centramos el contenido de la celda
            cabezera.Border.LeftBorder.BorderStyle = BorderStyleValues.Medium;
            cabezera.Border.RightBorder.BorderStyle = BorderStyleValues.Medium;
            cabezera.Border.BottomBorder.BorderStyle = BorderStyleValues.Medium;
            cabezera.Border.TopBorder.BorderStyle = BorderStyleValues.Medium;
            cabezera.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(255, 156, 193, 231), System.Drawing.Color.FromArgb(255, 156, 193, 231));
            //CABEZERA
            //fila 1
            reporte.SetCellStyle(1,1,1,numColDoc, cabezera);//damos el estilo a la celda A1
            reporte.SetCellValue("A1", "DISTRITO JUDICIAL DE MADRE DE DIOS\nEVOLUCION PERIODO\nFISCALIA ESPECIALIZADA EN CORRUPCION DE FUNCIONARIOS");
            reporte.MergeWorksheetCells(1, 1, 1, numColDoc);//mezclamos las celdas
            reporte.SetRowHeight(1, 45);//damos un alto de 45 a la fila 1
            //medidas columnas
            reporte.SetColumnWidth(1, 7);//damos un ancho a la columna 1
            reporte.SetColumnWidth(2, 15.5);//damos un ancho a la columna 1
            reporte.SetColumnWidth(3, 27);//damos un ancho a la columna 1
            //--estilo cuerpo parte cabezera sin negrita--
            SLStyle cbzrSN = reporte.GetCellStyle("A1");
            cbzrSN.Font.Bold = false;
            //--estilo cuerpo parte cabezera sin negrita y alineado a la izquierda--
            SLStyle cbzrSNIz = reporte.GetCellStyle("A1");
            cbzrSNIz.Font.Bold = false;
            cbzrSNIz.SetHorizontalAlignment(HorizontalAlignmentValues.Left);//centramos el contenido de la celda
            //--estilo cuerpo parte izquierda--
            SLStyle blancoFecha = reporte.GetCellStyle("A1");
            blancoFecha.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.White, System.Drawing.Color.White);
            //--estilo cuerpo parte izquierda--
            SLStyle blancoIzq = reporte.GetCellStyle("A1");
            blancoIzq.Font.Bold = false;
            blancoIzq.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.White, System.Drawing.Color.White);
            //--estilo para celdas doradas claro--
            SLStyle clrGold = reporte.GetCellStyle("A1");
            clrGold.Font.Bold = false;
            clrGold.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(255, 255, 230, 154), System.Drawing.Color.FromArgb(255, 255, 230, 154));
            //--estilo para celdas doradas claro 2--
            SLStyle clrGoldTwo = reporte.GetCellStyle("A1");
            clrGoldTwo.FormatCode = "#,##0.00";
            clrGoldTwo.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(255, 255, 230, 154), System.Drawing.Color.FromArgb(255, 255, 230, 154));
            //--estilo para celdas doradas claro porcentaje--
            SLStyle clrGoldPer = reporte.GetCellStyle("A1");
            clrGoldPer.Font.Bold = false;
            clrGoldPer.FormatCode = "0.00%";
            clrGoldPer.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(255, 255, 230, 154), System.Drawing.Color.FromArgb(255, 255, 230, 154));
            //--estilo para celdas purpura oscuro--
            SLStyle drkPrlp = reporte.GetCellStyle("A1");
            drkPrlp.Font.Bold = false;
            drkPrlp.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(255, 175, 169, 170), System.Drawing.Color.FromArgb(255, 175, 169, 170));
            //--estilo para celdas azul cyan--
            SLStyle blCyan = reporte.GetCellStyle("A1");
            blCyan.Font.Bold = false;
            blCyan.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(255, 172, 184, 203), System.Drawing.Color.FromArgb(255, 172, 184, 203));
            //--estilo para celdas amarillas--
            SLStyle yellow = reporte.GetCellStyle("A1");
            yellow.Font.Bold = false;
            yellow.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Yellow, System.Drawing.Color.Yellow);
            //--estilo para celdas amarillas porcentaje--
            SLStyle yellowPer = reporte.GetCellStyle("A1");
            yellowPer.Font.Bold = false;
            yellowPer.FormatCode = "0.00%";
            yellowPer.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Yellow, System.Drawing.Color.Yellow);
            //--estilo para celdas verde claro--
            SLStyle lghtGreen = reporte.GetCellStyle("A1");
            lghtGreen.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.LightGreen, System.Drawing.Color.LightGreen);
            //--estilo para mes inclinado
            SLStyle angle = reporte.GetCellStyle("A1");
            angle.Alignment.TextRotation = 75;
            //fila 2
            reporte.SetCellStyle(2, 1, 2, numColDoc, blancoFecha);//damos el estilo a la celda A2
            reporte.SetCellValue("A2", "PERIODO " + " A ");
            reporte.MergeWorksheetCells(2, 1, 2, numColDoc);//mezclamos las celdas
            reporte.SetRowHeight(2, 22.5);//damos un alto a la fila
            //CUERPO
            //fila 4
            reporte.SetCellStyle(4, 1, 4, 3, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("A4", "Inicio de vigencia NCPP : 01-10-2009");
            reporte.MergeWorksheetCells(4, 1, 4, 3);//mezclamos las celdas
            reporte.SetRowHeight(4, 60);//damos un alto a la fila
            //fila 5
            reporte.SetCellStyle(5, 1, 5, 3, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("A5", "FISCALES");
            reporte.MergeWorksheetCells(5, 1, 5, 3);//mezclamos las celdas
            reporte.SetRowHeight(5, 22.5);//damos un alto a la fila
            //fila 6
            reporte.SetRowHeight(6, 5);//damos un alto a la fila
            //fila 7
            reporte.SetCellStyle(7, 1, 7, 3, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("A7", "INGRESOS - ASIGNADOS");
            reporte.MergeWorksheetCells(7, 1, 7, 3);//mezclamos las celdas
            reporte.SetRowHeight(7, 22.5);//damos un alto a la fila
            //fila 8
            reporte.SetRowHeight(8, 5);//damos un alto a la fila
            //fila 9
            reporte.SetCellStyle(9, 1, 9, 3, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("A9", "DENUNCIAS - SALIDAS");
            reporte.MergeWorksheetCells(9, 1, 9, 3);//mezclamos las celdas
            reporte.SetRowHeight(9, 22.5);//damos un alto a la fila
            //fila 10
            reporte.SetCellStyle(10, 1, 10, 3, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("A10", "DERIVADOS");
            reporte.MergeWorksheetCells(10, 1, 10, 3);//mezclamos las celdas
            reporte.SetRowHeight(10, 22.5);//damos un alto a la fila
            //fila 11 columna A
            reporte.SetCellStyle(11, 1, 20, 1, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("A11", "EN\nM.P.");
            reporte.MergeWorksheetCells(11, 1, 20, 1);//mezclamos las celdas
            reporte.SetRowHeight(11, 22.5);//damos un alto a la fila
            //fila 11 columna B
            reporte.SetCellStyle(11, 2, 13, 2, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("B11", "ARCHIVO");
            reporte.MergeWorksheetCells(11, 2, 13, 2);//mezclamos las celdas
            //fila 11 columna C
            reporte.SetCellStyle(11, 3, drkPrlp);//damos el estilo a la celda
            reporte.SetCellValue("C11", "Consentido");
            //fila 12 columna C
            reporte.SetCellStyle(12, 3, blCyan);//damos el estilo a la celda
            reporte.SetCellValue("C12", "Sin Consentir");
            reporte.SetRowHeight(12, 22.5);//damos un alto a la fila
            //fila 13 columna C
            reporte.SetCellStyle(13, 3, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("C13", "TOTAL");
            reporte.SetRowHeight(13, 22.5);//damos un alto a la fila
            //fila 14 columns B
            reporte.SetCellStyle(14, 2, 16, 2, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("B14", "Vias Alternativas");
            reporte.MergeWorksheetCells(14, 2, 16, 2);//mezclamos las celdas
            reporte.SetRowHeight(14, 22.5);//damos un alto a la fila
            //fila 14 columns C
            reporte.SetCellStyle(14, 3, blCyan);//damos el estilo a la celda
            reporte.SetCellValue("C14", "Acuerdo Reparatorio");
            //fila 15 columna C
            reporte.SetCellStyle(15, 3, drkPrlp);//damos el estilo a la celda
            reporte.SetCellValue("C15", "Principio de Oportunidad");
            reporte.SetRowHeight(15, 22.5);//damos un alto a la fila
            //fila 16 columna C
            reporte.SetCellStyle(16, 3, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("C16", "TOTAL");
            reporte.SetRowHeight(16, 22.5);//damos un alto a la fila
            //fila 17 columna C
            reporte.SetCellStyle(17, 3, drkPrlp);//damos el estilo a la celda
            reporte.SetCellValue("C17", "Acumulado");
            reporte.SetRowHeight(17, 22.5);//damos un alto a la fila
            //fila 18 columna C
            reporte.SetCellStyle(18, 3, cbzrSN);//damos el estilo a la celda
            reporte.SetCellValue("C18", "Reserva Provisional");
            reporte.SetRowHeight(18, 22.5);//damos un alto a la fila
            //fila 19 columna C
            reporte.SetCellStyle(19, 3, drkPrlp);//damos el estilo a la celda
            reporte.SetCellValue("C19", "Reserva PNP");
            reporte.SetRowHeight(19, 22.5);//damos un alto a la fila
            //fila 20 columna B
            reporte.SetCellStyle(20, 2, 20, 3, yellow);//damos el estilo a la celda
            reporte.SetCellValue("B20", "TOTAL EN M.P.");
            reporte.MergeWorksheetCells(20, 2, 20, 3);//mezclamos las celdas
            reporte.SetRowHeight(20, 22.5);//damos un alto a la fila
            //fila 21 columna A
            reporte.SetCellStyle(21, 1, 26, 1, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("A21", "ANTE EL\nP.J.");
            reporte.MergeWorksheetCells(21, 1, 26, 1);//mezclamos las celdas
            reporte.SetRowHeight(21, 22.5);//damos un alto a la fila
            //fila 21 columna B
            reporte.SetCellStyle(21, 2, 23, 2, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("B21", "Sobreseimiento");
            reporte.MergeWorksheetCells(21, 2, 23, 2);//mezclamos las celdas
            //fila 21 columna C
            reporte.SetCellStyle(21, 3, drkPrlp);//damos el estilo a la celda
            reporte.SetCellValue("C21", "Consentido");
            //fila 22 columna C
            reporte.SetCellStyle(22, 3, cbzrSN);//damos el estilo a la celda
            reporte.SetCellValue("C22", "Sin Consentir");
            reporte.SetRowHeight(22, 22.5);//damos un alto a la fila
            //fila 23 columna C
            reporte.SetCellStyle(23, 3, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("C23", "TOTAL");
            reporte.SetRowHeight(23, 22.5);//damos un alto a la fila
            //fila 24 columna B
            reporte.SetCellStyle(24, 2, 25, 2, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("B24", "Procesos  *Simplificados\n* Comun");
            reporte.MergeWorksheetCells(24, 2, 25, 2);//mezclamos las celdas
            reporte.SetRowHeight(24, 22.5);//damos un alto a la fila
            //fila 24 columna C
            reporte.SetCellStyle(24, 3, cbzrSN);//damos el estilo a la celda
            reporte.SetCellValue("C24", "Sentencias");
            //fila 25 columna C
            reporte.SetCellStyle(25, 3, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("C25", "TOTAL");
            reporte.SetRowHeight(25, 22.5);//damos un alto a la fila
            //fila 26 columna A
            reporte.SetCellStyle(26, 2, 26, 3, yellow);//damos el estilo a la celda
            reporte.SetCellValue("B26", "TOTAL ANTE EL P.J.");
            reporte.MergeWorksheetCells(26, 2, 26, 3);//mezclamos las celdas
            reporte.SetRowHeight(26, 22.5);//damos un alto a la fila
            //fila 27
            reporte.SetCellStyle(27, 1, 27, 3, drkPrlp);//damos el estilo a la celda
            reporte.SetCellValue("A27", "SALIDA/RESUELTOS");
            reporte.MergeWorksheetCells(27, 1, 27, 3);//mezclamos las celdas
            reporte.SetRowHeight(27, 22.5);//damos un alto a la fila
            //fila 28
            reporte.SetRowHeight(28, 5);//damos un alto a la fila
            //fila 29
            reporte.SetCellStyle(29, 1, 29, 3, cbzrSN);//damos el estilo a la celda
            reporte.SetCellValue("A29", "TRAMITE");
            reporte.MergeWorksheetCells(29, 1, 29, 3);//mezclamos las celdas
            reporte.SetRowHeight(29, 22.5);//damos un alto a la fila
            //fila 30
            reporte.SetCellStyle(30, 1, 30, 3, drkPrlp);//damos el estilo a la celda
            reporte.SetCellValue("A30", "ANTIGUO CODIGO");
            reporte.MergeWorksheetCells(30, 1, 30, 3);//mezclamos las celdas
            reporte.SetRowHeight(30, 22.5);//damos un alto a la fila
            //fila 31
            reporte.SetRowHeight(31, 5);//damos un alto a la fila
            //fila 32
            reporte.SetCellStyle(32, 1, 32, 3, cbzrSN);//damos el estilo a la celda
            reporte.SetCellValue("A32", "EFICACIA (INGRESOS/SALIDAS)");
            reporte.MergeWorksheetCells(32, 1, 32, 3);//mezclamos las celdas
            reporte.SetRowHeight(32, 22.5);//damos un alto a la fila
            //fila 33
            reporte.SetRowHeight(33, 5);//damos un alto a la fila
            //fila 34
            reporte.SetCellStyle(34, 1, 34, 3, drkPrlp);//damos el estilo a la celda
            reporte.SetCellValue("A34", "PRODUCTIVIDAD");
            reporte.MergeWorksheetCells(34, 1, 34, 3);//mezclamos las celdas
            reporte.SetRowHeight(34, 22.5);//damos un alto a la fila
            //fila 35
            reporte.SetCellStyle(35, 1, 35, 3, blancoIzq);//damos el estilo a la celda
            reporte.SetCellValue("A35", "(Casos Fiscal / Mes) Terminados");
            reporte.MergeWorksheetCells(35, 1, 35, 3);//mezclamos las celdas
            reporte.SetRowHeight(35, 22.5);//damos un alto a la fila
            //fila 36
            reporte.SetCellStyle(36, 1, 36, numColDoc, cbzrSNIz);//damos el estilo a la celda A1
            reporte.SetCellValue("A36", "NOTA: La carga de los Fiscales es Obtenido de los Reportes del SGF (Reporte detallado de carga laboral)\n                El total de La Fiscalia es Obrenida del SFG (Reporte de Actos Procesales por Etapa)");
            reporte.MergeWorksheetCells(36, 1, 36, numColDoc);//mezclamos las celdas
            reporte.SetRowHeight(36, 33);//damos un alto de 45 a la fila 1
            int contador = 4;
            foreach (Fiscal fisc in Fiscalia)
            {
                double salRes = fisc.Deriv + fisc.Archicon + fisc.Archi + fisc.Acuerep + fisc.Princoport + fisc.Acum + fisc.Prov + fisc.Rpnp + fisc.Sobreseimcon + fisc.Sobreseim + fisc.Senten;
                //medidas columnas
                reporte.SetColumnWidth(contador, 11);//damos un ancho a la columna 1
                reporte.SetColumnWidth(contador + 1, 10);//damos un ancho a la columna 1
                reporte.SetColumnWidth(contador + 2, 3.7);//damos un ancho a la columna 1
                reporte.SetColumnWidth(contador + 3, 3.7);//damos un ancho a la columna 1
                //fila 4
                reporte.SetCellStyle(4, contador, 4, (contador + 3), blancoFecha);//damos el estilo a la celda
                reporte.SetCellValue(4, contador, fisc.Apell + ", " + fisc.Nomb);
                reporte.MergeWorksheetCells(4, contador, 4, (contador + 3));//mezclamos las celdas
                //fila 5
                reporte.SetCellStyle(5, contador, 5, (contador + 1), clrGold);//damos el estilo a la celda
                reporte.SetCellStyle(5, (contador + 2), 7, (contador + 3), blancoFecha);//damos el estilo a la celda
                reporte.SetCellValue(5, contador, 1);
                reporte.SetCellValue(5, contador + 2, "T");
                reporte.MergeWorksheetCells(5, contador, 5, (contador + 1));//mezclamos las celdas
                reporte.MergeWorksheetCells(5, (contador + 2), 7, (contador + 3));//mezclamos las celdas
                //fila 6
                reporte.SetCellStyle(6, contador, 6, (contador + 1), clrGold);//damos el estilo a la celda
                reporte.MergeWorksheetCells(6, contador, 6, (contador + 1));//mezclamos las celdas
                //fila 7
                reporte.SetCellStyle(7, contador, 7, (contador + 1), lghtGreen);//damos el estilo a la celda
                reporte.SetCellValue(7, contador, fisc.Asig);
                reporte.MergeWorksheetCells(7, contador, 7, (contador + 1));//mezclamos las celdas
                totalFiscal.Asig += fisc.Asig;
                //fila 8
                reporte.SetCellStyle(8, contador, 8, (contador + 1), clrGold);//damos el estilo a la celda
                reporte.MergeWorksheetCells(8, contador, 8, (contador + 1));//mezclamos las celdas
                //fila 8
                reporte.SetCellStyle(8, (contador + 2), 26, (contador + 3), angle);//damos el estilo a la celda
                reporte.MergeWorksheetCells(8, (contador + 2), 26, (contador + 3));//mezclamos las celdas
                //fila 9 primera columna
                reporte.SetCellStyle(9, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(9, contador, "Cantidad");
                //fila 9 segunda columna
                reporte.SetCellStyle(9, contador + 1, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(9, contador + 1, "%");
                //fila 10 primera columna
                reporte.SetCellStyle(10, contador, yellow);//damos el estilo a la celda
                reporte.SetCellValue(10, contador, fisc.Deriv);
                totalFiscal.Deriv += fisc.Deriv;
                //fila 10 segunda columna
                reporte.SetCellStyle(10, contador + 1, yellowPer);//damos el estilo a la celda
                reporte.SetCellValue(10, contador + 1, MayorZero(fisc.Deriv, salRes));
                //fila 11 primera columna
                reporte.SetCellStyle(11, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(11, contador, fisc.Archicon);
                totalFiscal.Archicon += fisc.Archicon;
                //fila 11 segunda columna
                reporte.SetCellStyle(11, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(11, contador + 1, MayorZero(fisc.Archicon, salRes));
                //fila 12 primera columna
                reporte.SetCellStyle(12, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(12, contador, fisc.Archi);
                totalFiscal.Archi += fisc.Archi;
                //fila 12 segunda columna
                reporte.SetCellStyle(12, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(12, contador + 1, MayorZero(fisc.Archi, salRes));
                //fila 13 primera columna
                reporte.SetCellStyle(13, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(13, contador, (fisc.Archicon + fisc.Archi));
                //fila 13 segunda columna
                reporte.SetCellStyle(13, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(13, contador + 1, MayorZero((fisc.Archicon + fisc.Archi), salRes));
                //fila 14 primera columna
                reporte.SetCellStyle(14, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(14, contador, fisc.Acuerep);
                totalFiscal.Acuerep +=fisc.Acuerep;
                //fila 14 segunda columna
                reporte.SetCellStyle(14, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(14, contador + 1, MayorZero(fisc.Acuerep, salRes));
                //fila 15 primera columna
                reporte.SetCellStyle(15, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(15, contador, fisc.Princoport);
                totalFiscal.Princoport +=fisc.Princoport;
                //fila 15 segunda columna
                reporte.SetCellStyle(15, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(15, contador + 1, MayorZero(fisc.Princoport, salRes));
                //fila 16 primera columna
                reporte.SetCellStyle(16, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(16, contador, (fisc.Acuerep + fisc.Princoport));
                //fila 16 segunda columna
                reporte.SetCellStyle(16, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(16, contador + 1, MayorZero((fisc.Acuerep + fisc.Princoport), salRes));
                //fila 17 primera columna
                reporte.SetCellStyle(17, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(17, contador, fisc.Acum);
                totalFiscal.Acum += fisc.Acum;
                //fila 17 segunda columna
                reporte.SetCellStyle(17, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(17, contador + 1, MayorZero(fisc.Acum, salRes));
                //fila 18 primera columna
                reporte.SetCellStyle(18, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(18, contador, fisc.Prov);
                totalFiscal.Prov += fisc.Prov;
                //fila 18 segunda columna
                reporte.SetCellStyle(18, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(18, contador + 1, MayorZero(fisc.Prov, salRes));
                //fila 19 primera columna
                reporte.SetCellStyle(19, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(19, contador, fisc.Rpnp);
                totalFiscal.Rpnp +=fisc.Rpnp;
                //fila 19 segunda columna
                reporte.SetCellStyle(19, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(19, contador + 1, MayorZero(fisc.Rpnp, salRes));
                //fila 20 primera columna
                reporte.SetCellStyle(20, contador, yellow);//damos el estilo a la celda
                reporte.SetCellValue(20, contador, (fisc.Deriv + fisc.Archicon + fisc.Archi + fisc.Acuerep + fisc.Princoport + fisc.Acum + fisc.Prov + fisc.Rpnp));
                //fila 20 segunda columna
                reporte.SetCellStyle(20, contador + 1, yellowPer);//damos el estilo a la celda
                reporte.SetCellValue(20, contador + 1, MayorZero((fisc.Deriv + fisc.Archicon + fisc.Archi + fisc.Acuerep + fisc.Princoport + fisc.Acum + fisc.Prov + fisc.Rpnp), salRes));
                //fila 21 primera columna
                reporte.SetCellStyle(21, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(21, contador, fisc.Sobreseimcon);
                totalFiscal.Sobreseimcon +=fisc.Sobreseimcon;
                //fila 21 segunda columna
                reporte.SetCellStyle(21, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(21, contador + 1, MayorZero(fisc.Sobreseimcon, salRes));
                //fila 22 primera columna
                reporte.SetCellStyle(22, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(22, contador, fisc.Sobreseim);
                totalFiscal.Sobreseim += fisc.Sobreseim;
                //fila 22 segunda columna
                reporte.SetCellStyle(22, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(22, contador + 1, MayorZero(fisc.Sobreseim, salRes));
                //fila 23 primera columna
                reporte.SetCellStyle(23, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(23, contador, (fisc.Sobreseimcon+fisc.Sobreseim));
                //fila 23 segunda columna
                reporte.SetCellStyle(23, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(23, contador + 1, MayorZero((fisc.Sobreseimcon + fisc.Sobreseim), salRes));
                //fila 24 primera columna
                reporte.SetCellStyle(24, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(24, contador, fisc.Senten);
                totalFiscal.Senten += fisc.Senten;
                //fila 24 segunda columna
                reporte.SetCellStyle(24, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(24, contador + 1, MayorZero(fisc.Senten, salRes));
                //fila 25 primera columna
                reporte.SetCellStyle(25, contador, clrGold);//damos el estilo a la celda
                reporte.SetCellValue(25, contador, fisc.Senten);
                //fila 25 segunda columna
                reporte.SetCellStyle(25, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(25, contador + 1, MayorZero(fisc.Senten, salRes));
                //fila 26 primera columna
                reporte.SetCellStyle(26, contador, yellow);//damos el estilo a la celda
                reporte.SetCellValue(26, contador, (fisc.Sobreseimcon + fisc.Sobreseim + fisc.Senten));
                //fila 26 segunda columna
                reporte.SetCellStyle(26, contador + 1, yellowPer);//damos el estilo a la celda
                reporte.SetCellValue(26, contador + 1, MayorZero((fisc.Sobreseimcon + fisc.Sobreseim + fisc.Senten), salRes));
                //fila 27
                reporte.SetCellStyle(27, contador, 27, (contador + 1), clrGold);//damos el estilo a la celda
                reporte.SetCellValue(27, contador, salRes);
                reporte.MergeWorksheetCells(27, contador, 27, (contador + 1));//mezclamos las celdas
                //fila 28
                reporte.SetCellStyle(28, contador, 28, (contador + 1), clrGold);//damos el estilo a la celda
                reporte.MergeWorksheetCells(28, contador, 28, (contador + 1));//mezclamos las celdas
                //fila 29
                reporte.SetCellStyle(29, contador, 29, (contador + 1), lghtGreen);//damos el estilo a la celda
                reporte.SetCellValue(29, contador, (fisc.Asig-salRes)-(fisc.Expe+fisc.Iexpe+fisc.Ipre));
                reporte.MergeWorksheetCells(29, contador, 29, (contador + 1));//mezclamos las celdas
                //fila 30
                reporte.SetCellStyle(30, contador, 30, (contador + 1), clrGold);//damos el estilo a la celda
                reporte.SetCellValue(30, contador, (fisc.Expe + fisc.Iexpe + fisc.Ipre));
                reporte.MergeWorksheetCells(30, contador, 30, (contador + 1));//mezclamos las celdas
                totalFiscal.Expe +=fisc.Expe;
                totalFiscal.Iexpe += fisc.Iexpe;
                totalFiscal.Ipre += fisc.Ipre;
                //fila 31
                reporte.SetCellStyle(31, contador, 31, (contador + 1), clrGold);//damos el estilo a la celda
                reporte.MergeWorksheetCells(31, contador, 31, (contador + 1));//mezclamos las celdas
                //fila 32
                reporte.SetCellStyle(32, contador, 32, contador + 1, clrGoldPer);//damos el estilo a la celda
                reporte.SetCellValue(32, contador, MayorZero(salRes, fisc.Asig));
                reporte.MergeWorksheetCells(32, contador, 32, (contador + 1));//mezclamos las celdas
                //fila 33
                reporte.SetCellStyle(33, contador, 33, (contador + 1), clrGold);//damos el estilo a la celda
                reporte.MergeWorksheetCells(33, contador, 33, (contador + 1));//mezclamos las celdas
                //fila 34
                reporte.SetCellStyle(34, contador, 34, contador + 1, clrGoldTwo);//damos el estilo a la celda
                reporte.SetCellValue(34, contador, (salRes/1)/24);
                reporte.MergeWorksheetCells(34, contador, 34, (contador + 1));//mezclamos las celdas
                //fila 35
                reporte.SetCellStyle(35, contador, 35, (contador + 1), clrGold);//damos el estilo a la celda
                reporte.MergeWorksheetCells(35, contador, 35, (contador + 1));//mezclamos las celdas
                contador += 4;
            }
            double totalSalRes = totalFiscal.Deriv + totalFiscal.Archicon + totalFiscal.Archi + totalFiscal.Acuerep + totalFiscal.Princoport + totalFiscal.Acum + totalFiscal.Prov + totalFiscal.Rpnp + totalFiscal.Sobreseimcon + totalFiscal.Sobreseim + totalFiscal.Senten;
            //medidas columnas
            reporte.SetColumnWidth(numColDoc-3, 11);//damos un ancho a la columna 1
            reporte.SetColumnWidth(numColDoc - 2, 10);//damos un ancho a la columna 1
            reporte.SetColumnWidth(numColDoc - 1, 3.7);//damos un ancho a la columna 1
            reporte.SetColumnWidth(numColDoc, 3.7);//damos un ancho a la columna 1
            //fila 4
            reporte.SetCellStyle(4, numColDoc - 3, 4, numColDoc, blancoFecha);//damos el estilo a la celda
            reporte.SetCellValue(4,numColDoc-3, "CARGA TOTAL FISCALIA");
            reporte.MergeWorksheetCells(4, numColDoc - 3, 4, numColDoc);//mezclamos las celdas
            //--------------
            //fila 5
            reporte.SetCellStyle(5, contador, 5, (contador + 1), clrGold);//damos el estilo a la celda
            reporte.SetCellStyle(5, (contador + 2), 7, (contador + 3), blancoFecha);//damos el estilo a la celda
            reporte.SetCellValue(5, contador, Fiscalia.Count);
            reporte.MergeWorksheetCells(5, contador, 5, (contador + 1));//mezclamos las celdas
            reporte.MergeWorksheetCells(5, (contador + 2), 7, (contador + 3));//mezclamos las celdas
            //fila 6
            reporte.SetCellStyle(6, contador, 6, (contador + 1), clrGold);//damos el estilo a la celda
            reporte.MergeWorksheetCells(6, contador, 6, (contador + 1));//mezclamos las celdas
            //fila 7
            reporte.SetCellStyle(7, contador, 7, (contador + 1), lghtGreen);//damos el estilo a la celda
            reporte.SetCellValue(7, contador, totalFiscal.Asig);
            reporte.MergeWorksheetCells(7, contador, 7, (contador + 1));//mezclamos las celdas
            //fila 8
            reporte.SetCellStyle(8, contador, 8, (contador + 1), clrGold);//damos el estilo a la celda
            reporte.MergeWorksheetCells(8, contador, 8, (contador + 1));//mezclamos las celdas
            //fila 8
            reporte.SetCellStyle(8, (contador + 2), 26, (contador + 3), angle);//damos el estilo a la celda
            reporte.MergeWorksheetCells(8, (contador + 2), 26, (contador + 3));//mezclamos las celdas
            //fila 9 primera columna
            reporte.SetCellStyle(9, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(9, contador, "Cantidad");
            //fila 9 segunda columna
            reporte.SetCellStyle(9, contador + 1, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(9, contador + 1, "%");
            //fila 10 primera columna
            reporte.SetCellStyle(10, contador, yellow);//damos el estilo a la celda
            reporte.SetCellValue(10, contador, totalFiscal.Deriv);
            //fila 10 segunda columna
            reporte.SetCellStyle(10, contador + 1, yellowPer);//damos el estilo a la celda
            reporte.SetCellValue(10, contador + 1, MayorZero(totalFiscal.Deriv, totalSalRes));
            //fila 11 primera columna
            reporte.SetCellStyle(11, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(11, contador, totalFiscal.Archicon);
            //fila 11 segunda columna
            reporte.SetCellStyle(11, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(11, contador + 1, MayorZero(totalFiscal.Archicon, totalSalRes));
            //fila 12 primera columna
            reporte.SetCellStyle(12, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(12, contador, totalFiscal.Archi);
            //fila 12 segunda columna
            reporte.SetCellStyle(12, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(12, contador + 1, MayorZero(totalFiscal.Archi, totalSalRes));
            //fila 13 primera columna
            reporte.SetCellStyle(13, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(13, contador, (totalFiscal.Archicon + totalFiscal.Archi));
            //fila 13 segunda columna
            reporte.SetCellStyle(13, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(13, contador + 1, MayorZero((totalFiscal.Archicon + totalFiscal.Archi), totalSalRes));
            //fila 14 primera columna
            reporte.SetCellStyle(14, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(14, contador, totalFiscal.Acuerep);
            //fila 14 segunda columna
            reporte.SetCellStyle(14, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(14, contador + 1, MayorZero(totalFiscal.Acuerep, totalSalRes));
            //fila 15 primera columna
            reporte.SetCellStyle(15, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(15, contador, totalFiscal.Princoport);
            //fila 15 segunda columna
            reporte.SetCellStyle(15, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(15, contador + 1, MayorZero(totalFiscal.Princoport, totalSalRes));
            //fila 16 primera columna
            reporte.SetCellStyle(16, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(16, contador, (totalFiscal.Acuerep + totalFiscal.Princoport));
            //fila 16 segunda columna
            reporte.SetCellStyle(16, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(16, contador + 1, MayorZero((totalFiscal.Acuerep + totalFiscal.Princoport), totalSalRes));
            //fila 17 primera columna
            reporte.SetCellStyle(17, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(17, contador, totalFiscal.Acum);
            //fila 17 segunda columna
            reporte.SetCellStyle(17, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(17, contador + 1, MayorZero(totalFiscal.Acum, totalSalRes));
            //fila 18 primera columna
            reporte.SetCellStyle(18, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(18, contador, totalFiscal.Prov);
            //fila 18 segunda columna
            reporte.SetCellStyle(18, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(18, contador + 1, MayorZero(totalFiscal.Prov, totalSalRes));
            //fila 19 primera columna
            reporte.SetCellStyle(19, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(19, contador, totalFiscal.Rpnp);
            //fila 19 segunda columna
            reporte.SetCellStyle(19, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(19, contador + 1, MayorZero(totalFiscal.Rpnp, totalSalRes));
            //fila 20 primera columna
            reporte.SetCellStyle(20, contador, yellow);//damos el estilo a la celda
            reporte.SetCellValue(20, contador, (totalFiscal.Deriv + totalFiscal.Archicon + totalFiscal.Archi + totalFiscal.Acuerep + totalFiscal.Princoport + totalFiscal.Acum + totalFiscal.Prov + totalFiscal.Rpnp));
            //fila 20 segunda columna
            reporte.SetCellStyle(20, contador + 1, yellowPer);//damos el estilo a la celda
            reporte.SetCellValue(20, contador + 1, MayorZero((totalFiscal.Deriv + totalFiscal.Archicon + totalFiscal.Archi + totalFiscal.Acuerep + totalFiscal.Princoport + totalFiscal.Acum + totalFiscal.Prov + totalFiscal.Rpnp), totalSalRes));
            //fila 21 primera columna
            reporte.SetCellStyle(21, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(21, contador, totalFiscal.Sobreseimcon);
            //fila 21 segunda columna
            reporte.SetCellStyle(21, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(21, contador + 1, MayorZero(totalFiscal.Sobreseimcon, totalSalRes));
            //fila 22 primera columna
            reporte.SetCellStyle(22, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(22, contador, totalFiscal.Sobreseim);
            //fila 22 segunda columna
            reporte.SetCellStyle(22, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(22, contador + 1, MayorZero(totalFiscal.Sobreseim, totalSalRes));
            //fila 23 primera columna
            reporte.SetCellStyle(23, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(23, contador, (totalFiscal.Sobreseimcon + totalFiscal.Sobreseim));
            //fila 23 segunda columna
            reporte.SetCellStyle(23, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(23, contador + 1, MayorZero((totalFiscal.Sobreseimcon + totalFiscal.Sobreseim), totalSalRes));
            //fila 24 primera columna
            reporte.SetCellStyle(24, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(24, contador, totalFiscal.Senten);
            //fila 24 segunda columna
            reporte.SetCellStyle(24, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(24, contador + 1, MayorZero(totalFiscal.Senten, totalSalRes));
            //fila 25 primera columna
            reporte.SetCellStyle(25, contador, clrGold);//damos el estilo a la celda
            reporte.SetCellValue(25, contador, totalFiscal.Senten);
            //fila 25 segunda columna
            reporte.SetCellStyle(25, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(25, contador + 1, MayorZero(totalFiscal.Senten, totalSalRes));
            //fila 26 primera columna
            reporte.SetCellStyle(26, contador, yellow);//damos el estilo a la celda
            reporte.SetCellValue(26, contador, (totalFiscal.Sobreseimcon + totalFiscal.Sobreseim + totalFiscal.Senten));
            //fila 26 segunda columna
            reporte.SetCellStyle(26, contador + 1, yellowPer);//damos el estilo a la celda
            reporte.SetCellValue(26, contador + 1, MayorZero((totalFiscal.Sobreseimcon + totalFiscal.Sobreseim + totalFiscal.Senten), totalSalRes));
            //fila 27
            reporte.SetCellStyle(27, contador, 27, (contador + 1), clrGold);//damos el estilo a la celda
            reporte.SetCellValue(27, contador, totalSalRes);
            reporte.MergeWorksheetCells(27, contador, 27, (contador + 1));//mezclamos las celdas
            //fila 28
            reporte.SetCellStyle(28, contador, 28, (contador + 1), clrGold);//damos el estilo a la celda
            reporte.MergeWorksheetCells(28, contador, 28, (contador + 1));//mezclamos las celdas
            //fila 29
            reporte.SetCellStyle(29, contador, 29, (contador + 1), lghtGreen);//damos el estilo a la celda
            reporte.SetCellValue(29, contador, (totalFiscal.Asig - totalSalRes) - (totalFiscal.Expe + totalFiscal.Iexpe + totalFiscal.Ipre));
            reporte.MergeWorksheetCells(29, contador, 29, (contador + 1));//mezclamos las celdas
            //fila 30
            reporte.SetCellStyle(30, contador, 30, (contador + 1), clrGold);//damos el estilo a la celda
            reporte.SetCellValue(30, contador, (totalFiscal.Expe + totalFiscal.Iexpe + totalFiscal.Ipre));
            reporte.MergeWorksheetCells(30, contador, 30, (contador + 1));//mezclamos las celdas
            //fila 31
            reporte.SetCellStyle(31, contador, 31, (contador + 1), clrGold);//damos el estilo a la celda
            reporte.MergeWorksheetCells(31, contador, 31, (contador + 1));//mezclamos las celdas
            //fila 32
            reporte.SetCellStyle(32, contador, 32, contador + 1, clrGoldPer);//damos el estilo a la celda
            reporte.SetCellValue(32, contador, MayorZero(totalSalRes, totalFiscal.Asig));
            reporte.MergeWorksheetCells(32, contador, 32, (contador + 1));//mezclamos las celdas
            //fila 33
            reporte.SetCellStyle(33, contador, 33, (contador + 1), clrGold);//damos el estilo a la celda
            reporte.MergeWorksheetCells(33, contador, 33, (contador + 1));//mezclamos las celdas
            //fila 34
            reporte.SetCellStyle(34, contador, 34, contador + 1, clrGoldTwo);//damos el estilo a la celda
            reporte.SetCellValue(34, contador, (totalSalRes / Fiscalia.Count)/24);
            reporte.MergeWorksheetCells(34, contador, 34, (contador + 1));//mezclamos las celdas
            //fila 35
            reporte.SetCellStyle(35, contador, 35, (contador + 1), clrGold);//damos el estilo a la celda
            reporte.MergeWorksheetCells(35, contador, 35, (contador + 1));//mezclamos las celdas
            contador += 4;
            //guardando documento
            reporte.SaveAs("prueba01.xlsx");//guardamos el documento
        }

        public static double MayorZero (double val1, double val2)
        {
            if (val1 > 0 && val2 > 0)
            {
                return (val1 / val2);
            }
            else
            {
                return 0;
            }
        }
    }
}
