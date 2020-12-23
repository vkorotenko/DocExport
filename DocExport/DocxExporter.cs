using System.Data;
using System.IO;
using NPOI.XWPF.UserModel;

namespace DocExport
{
    public class DocxExporter
    {
        private const string Fio = "FIO";
        private const string IdOrders = "id_orders";
        private const string NName = "Nname";
        private const string IdStaff = "idstaff";
        private const string Position = "Position";

        /// <summary>
        /// Экспортируем подготовленный датасет в ворд
        /// </summary>
        /// <param name="path">Путь по которому создается файл</param>
        /// <param name="dataSet">Датасет с данными</param>
        public static void Export(string path, DataSet dataSet)
        {
            var doc = new XWPFDocument();
            var para = doc.CreateParagraph();
            var r0 = para.CreateRun();
            r0.SetText("Таблица заказов");
            para.FillBackgroundColor = "EEEEEE";
            // para.FillPattern = NPOI.OpenXmlFormats.Wordprocessing.ST_Shd.diagStripe;

            var data = dataSet.Tables[0];
            var table = doc.CreateTable(data.Rows.Count, 5);
            var pos = 0;
            foreach (DataRow row in data.Rows)
            {
                CreateRow(table, row, pos++);
            }
            var outStream = new FileStream(path, FileMode.Create);
            doc.Write(outStream);
            outStream.Close();
        }

        private static void CreateRow(XWPFTable table, DataRow row , int pos)
        {
            table.GetRow(pos).GetCell(0).SetText(row[Fio].ToString());
            table.GetRow(pos).GetCell(1).SetText(row[IdOrders].ToString());
            table.GetRow(pos).GetCell(2).SetText(row[NName].ToString());
            table.GetRow(pos).GetCell(3).SetText(row[IdStaff].ToString());
            table.GetRow(pos).GetCell(4).SetText(row[Position].ToString());

        }
    }
}
