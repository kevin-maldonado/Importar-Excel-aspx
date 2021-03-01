using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System.Data;
using System.Data.OleDb;

namespace WebEjemplo
{
    public partial class Excel : System.Web.UI.Page
    {
        public static string path = @"C:\Users\kevin\OneDrive\Escritorio\ProyectoExcel\miexcel.xlsx";
        public static string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extend Properties=Excel 12.0;";




        protected void Page_Load(object sender, EventArgs e, string strm)
        {

        }
        //metodo para cargar los datos 
        private void CargarDatos(string strm)
        {
            SLDocument sl = new SLDocument(path);
            //desde donde se va a inicar los datos
            int iRow = 2;
            //se crea la lista view models
            List<LogicaExcel> lst = new List<LogicaExcel>();

            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
            {
                LogicaExcel Usuario = new LogicaExcel();
                Usuario.codigo = sl.GetCellValueAsString(iRow, 1);
                Usuario.nombre = sl.GetCellValueAsString(iRow, 2);
                Usuario.edad = sl.GetCellValueAsInt32(iRow, 3);
                lst.Add(Usuario);

                iRow++;

            }
            //se muestra los datos en el data grid view

            GridView1.DataSource = lst;
            GridView1.DataBind();
        }
        ////metodo para guardar la tabla en base de datos
        //private DataTable CrearTabla(String fila)
        //{
        //    int cantidadColumnas;
        //    DataTable tabla = new DataTable("miexcel");
        //    String[] valores = fila.Split(new char[] { ',' });
        //    cantidadColumnas = valores.Length;
        //    int idx = 0;
        //    foreach (String val in valores)
        //    {
        //        String nombreColumna = String.Format("{0}", idx++);
        //        tabla.Columns.Add(nombreColumna, Type.GetType("System.String"));
        //    }
        //    return tabla;
        //}
        
        //metodo para verificar el tipo de archivo
        bool ChecarExtension(string fileName)
        {
            string ext = Path.GetExtension(fileName);
            switch (ext.ToLower())
            {
                case ".xlsx":
                    return true;
                default:
                    return false;
            }
        }
        //botones
        //para subir el archivo de excel
        public void Button1_Click(object sender, EventArgs e)
        {
            
            if (FileUpload1.HasFile)
            {
                if (ChecarExtension(FileUpload1.FileName))
                {
                    
                    FileUpload1.SaveAs(MapPath("~/ArchivoExcel/" + FileUpload1.FileName));



                    Label1.Text = FileUpload1.FileName + " archivo cargado exitosamente";

                    lblOculto.Text = MapPath("~/ArchivoExcel/" + FileUpload1.FileName);
                }
            }
            else
            {
                Label1.Text = "Error al subir el archivo excel";
            }

        }
        //boton 2 carga datos
        protected void Button2_Click(object sender, EventArgs e)
        {
            try
            {

                CargarDatos(lblOculto.Text);
            }
            catch
            {
                Response.Write("Ocurrio un error debe cargar antes el archivo");
            }
        }

        protected void Button3_Click(object sender, EventArgs e)
        {
            string path = @"C:\Users\kevin\OneDrive\Escritorio\ProyectoExcel\miexcel.xlsx";
            SLDocument sl = new SLDocument(path);
            using (var db = new pruebaEntities1())
            {

                int iRow = 2;
                while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
                {
                    string codigo = sl.GetCellValueAsString(iRow, 1);
                    string nombre = sl.GetCellValueAsString(iRow, 2);
                    int edad = sl.GetCellValueAsInt32(iRow, 3);

                    var oMiExcel = new miexcel();
                    oMiExcel.codigo = codigo;
                    oMiExcel.nombre = nombre;
                    oMiExcel.edad = edad;

                    db.miexcel.Add(oMiExcel);
                    db.SaveChanges();

                    iRow++;
                }
            }
        }
    }
}