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
        public static string path = @"C:\Users\kevin\OneDrive\Escritorio\Importar-Excel-aspx\archivoexcelparasubir\miexcel.xlsx";//aqui va la direccion del documento excel que se va a subir
        public static string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extend Properties=Excel 12.0;";
        //Cargar los datos del excel
        protected void CargarDatos(string strm)
        {
            SLDocument sl = new SLDocument(path);
            //desde donde se va a inicar los datos
            int iRow = 2;
            //se crean las lista LogicaExcel
            List<LogicaExcel> list1 = new List<LogicaExcel>();
            List<LogicaExcel> list2 = new List<LogicaExcel>();
            //Valida si los campos del excel estan sin ningun codigo
            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
            {
                LogicaExcel Usuario = new LogicaExcel();
                Usuario.codigo = sl.GetCellValueAsInt32(iRow, 1);
                Usuario.nombre = sl.GetCellValueAsString(iRow, 2);
                Usuario.edad = sl.GetCellValueAsInt32(iRow, 3);
                //Condicion para verificar los campos vacios del nombre
                if (Usuario.nombre != String.Empty)
                {
                    list1.Add(Usuario);
                }
                else 
                {
                    list2.Add(Usuario);
                }
                iRow++;
            }
            //condicin para mostrar campos vacios en el gridview
            if (list2.Count() != 0)  
            {
                //se muestra los datos en el grid view
                GridView1.DataSource = list2;
                GridView1.DataBind();
            }
            else 
            {
                //se muestra los datos en el grid view
                GridView1.DataSource = list1;
                GridView1.DataBind();
            }  
        }
      
        protected void Page_Load(object sender, EventArgs e, string strm)
        {
            if (!IsPostBack)
            {
      
            }

        }
       
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
        //para subir el archivo de excel y cargar los datos
        public void Subir_Onclick(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile)
            {
                if (ChecarExtension(FileUpload1.FileName))
                {
                    FileUpload1.SaveAs(MapPath("~/ArchivoExcel/" + FileUpload1.FileName));
                    Lbl_Mensaje.Text = FileUpload1.FileName + " archivo cargado exitosamente";
                    Lbl_Oculto.Text = MapPath("~/ArchivoExcel/" + FileUpload1.FileName);
                    try
                    {
                        CargarDatos(Lbl_Oculto.Text);
                    }
                    catch
                    {
                        Response.Write("Ocurrio un error debe cargar antes el archivo");
                    }
                }
            }
            else
            {
                Lbl_Mensaje.Text = "Error al subir el archivo excel";
            }
        }
        //Dar Colores las filas que estan vacias al cargar el excel
        protected void grdDatos_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                string nombre = DataBinder.Eval(e.Row.DataItem, "nombre").ToString();

                if (nombre != String.Empty)
                {
                    e.Row.BackColor = System.Drawing.Color.Transparent;
                   
                }
                else
                {
                    Btn_Guardar.Visible = false;
                    e.Row.BackColor = System.Drawing.Color.Red;
                    Lbl_Mensaje.Text = "Llenar todos los campos vacios de la tabla por favor";
                }           
            }
        }
        protected void Guardar_Onclick(object sender, EventArgs e)
        {
            string path = @"C:\Users\kevin\OneDrive\Escritorio\Importar-Excel-aspx\archivoexcelparasubir\miexcel.xlsx";//direccion del archivos para que cargue los datos en la base
            SLDocument sl = new SLDocument(path);
            using (var db = new PruebaEntities())
            {
                int iRow = 2;
                while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
                {
                    int codigo = sl.GetCellValueAsInt32(iRow, 1);
                    string nombre = sl.GetCellValueAsString(iRow, 2);
                    int edad = sl.GetCellValueAsInt32(iRow, 3);
                    var Persona = new Persona();
                    Persona.codigo = codigo;
                    Persona.nombre = nombre;
                    Persona.edad = edad;
                    db.Persona.Add(Persona);
                    db.SaveChanges();
                    iRow++;
                    Lbl_Mensaje.Text = "Datos Guardados Exitosamente";
                }
            }
        }
    }
}