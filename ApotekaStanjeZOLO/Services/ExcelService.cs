using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ApotekaStanjeZOLO.Services
{
    class ExcelService
    {
        public List<string> VratiListuZolo(string excel)
        {
            var lista = new List<string>();

            var wb = new XLWorkbook(excel);

            var ws = wb.Worksheet(1);

            var firstRowUsed = ws.FirstRowUsed();


            while (!firstRowUsed.Cell(1).IsEmpty())
            {

                string JKL = firstRowUsed.Cell(1).GetString();

                lista.Add(JKL);

                firstRowUsed = firstRowUsed.RowBelow();
            }
            return lista;
        }
    }
}
