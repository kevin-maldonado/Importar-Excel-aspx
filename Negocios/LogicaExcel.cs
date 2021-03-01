using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace Negocios
{
    public class LogicaExcel1 : DataRow
    {
        public int id
        {
            get { return (int)base["id"]; }
            set { base["id"] = value; }
        }
        public string codigo
        {
            get { return (string)base["codigo"]; }
            set { base["codigo"] = value; }
        }
        public string nombre
        {
            get { return (string)base["nombre"]; }
            set { base["nombre"] = value; }
        }
        public int edad
        {
            get { return (int)base["edad"]; }
            set { base["edad"] = value; }
        }

        internal LogicaExcel1(DataRowBuilder builder)
            : base(builder)
        {
            this.id = 0;
            this.codigo = string.Empty;
            this.nombre = string.Empty;
            this.edad = 0;
        }
    }

    public class ProductTable : DataTable
    {
        public ProductTable(string miexcel)
        {
            this.TableName = TableName;
            this.Columns.Add(new DataColumn("id", typeof(int)));
            this.Columns.Add(new DataColumn("codigo", typeof(string)));
            this.Columns.Add(new DataColumn("DateAdded", typeof(string)));
            this.Columns.Add(new DataColumn("edad", typeof(int)));
        }

        public LogicaExcel1 CreateNewRow()
        {
            return (LogicaExcel1)NewRow();
        }

        protected override Type GetRowType()
        {
            return typeof(LogicaExcel1);
        }

        protected override DataRow NewRowFromBuilder(DataRowBuilder builder)
        {
            return new LogicaExcel1(builder);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            DataTable dtStrongTyping = new DataTable("miexcel");
            ProductTable dtSchwarzeneggerTyping = new ProductTable("ProductSchwarzeneggerTyping");

            // this basically makes it equivalent to ProductTable,
            // but we're "manually" doing it in code, instead of having
            // all this done by a "proper" class.
            dtStrongTyping.Columns.Add(new DataColumn("id", typeof(int)));
            dtStrongTyping.Columns.Add(new DataColumn("codigo", typeof(string)));
            dtStrongTyping.Columns.Add(new DataColumn("nombre", typeof(string)));
            dtStrongTyping.Columns.Add(new DataColumn("edad", typeof(int)));

            int row;
            LogicaExcel1 prodrow;

            using (SLDocument sl = new SLDocument("miexcel.xlsx", "Hoja1"))
            {
                SLWorksheetStatistics stats = sl.GetWorksheetStatistics();
                int iStartColumnIndex = stats.StartColumnIndex;

                // I'll assume that the "first" row is always the header row.
                // Notice that the StartRowIndex returned isn't necessarily the
                // first row of the worksheet, it's the first row that has any data.

                // WARNING: the statistics object notes down any non-empty cells.
                // This includes cells with say a background colour but doesn't have
                // cell data. So if you have an empty row just after your worksheet
                // tabular data, but the row is coloured light blue, the EndRowIndex
                // will be one more than you need.
                // It is suggested that you know more about the input Excel file you're
                // using...

                // I'm also not using any variables to store the intermediate returned
                // cell data. This makes each code segment independent of each other,
                // and also makes it such that it's easier for you to see what you
                // actually have to type.

                for (row = stats.StartRowIndex + 1; row <= stats.EndRowIndex; ++row)
                {
                    dtStrongTyping.Rows.Add(
                        sl.GetCellValueAsInt32(row, iStartColumnIndex),
                        sl.GetCellValueAsString(row, iStartColumnIndex + 1),
                        sl.GetCellValueAsString(row, iStartColumnIndex + 2),
                        sl.GetCellValueAsInt32(row, iStartColumnIndex + 3)
                        );
                }


                for (row = stats.StartRowIndex + 1; row <= stats.EndRowIndex; ++row)
                {
                    prodrow = dtSchwarzeneggerTyping.CreateNewRow();
                    prodrow.id = sl.GetCellValueAsInt32(row, iStartColumnIndex);
                    prodrow.nombre = sl.GetCellValueAsString(row, iStartColumnIndex + 1);
                    prodrow.codigo = sl.GetCellValueAsString(row, iStartColumnIndex + 2);
                    prodrow.edad = sl.GetCellValueAsInt32(row, iStartColumnIndex + 3);
                    dtSchwarzeneggerTyping.Rows.Add(prodrow);
                }
            }

            // just to prove that the data in each DataTable is correct...
            dtStrongTyping.Rows.Add(1, "I change keyboards every month because they can't handle my strong typing skills", DateTime.Now.AddMonths(1), 2.78m);

            prodrow = dtSchwarzeneggerTyping.CreateNewRow();
            prodrow.id = 1;
            prodrow.codigo = "SAMN";
            prodrow.nombre = "kevin";
            prodrow.edad = 22;
            dtSchwarzeneggerTyping.Rows.Add(prodrow);

            // I don't know what you're going to do with the DataTable,
            // so I'm just going to write the contents to XML files.

            dtStrongTyping.WriteXml("ExportToDataTableExampleStrongTyping.xml");
            dtSchwarzeneggerTyping.WriteXml("ExportToDataTableExampleSchwarzeneggerTyping.xml");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
