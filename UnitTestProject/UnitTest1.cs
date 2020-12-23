using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;
using System.Diagnostics.PerformanceData;
using DocExport;

namespace UnitTestProject
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            const string path = @"c:\temp\test.docx";
            var dataSet = new DataSet();
            var table = new DataTable("users");

            #region Создание столбцов
            table.Columns.Add("FIO");
            table.Columns.Add("id_orders");
            table.Columns.Add("Nname");
            table.Columns.Add("idstaff");
            table.Columns.Add("Position");
            #endregion
            #region Заполнение данными
            table.Rows.Add("Иванов Иван Иванович", "5", "nname 1", "idstaff 1", "Position 1");
            table.Rows.Add("Петров Иван Иванович", "6", "nname 2", "idstaff 2", "Position 2");
            table.Rows.Add("Сидоров Иван Иванович", "7", "nname 3", "idstaff 3", "Position 3");
            table.Rows.Add("Петрыкин Иван Иванович", "8", "nname 4", "idstaff 4", "Position 4");
            #endregion

            dataSet.Tables.Add(table);
            DocxExporter.Export(path, dataSet);
        }
    }
}
