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
            //se crea la lista view models
            List<LogicaExcel> list1 = new List<LogicaExcel>();
            List<LogicaExcel> list2 = new List<LogicaExcel>();
            //Valida si los campos del excel estan sin ningun codigo
            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iRow, 1)))
            {
                LogicaExcel Usuario = new LogicaExcel();
                Usuario.codigo = sl.GetCellValueAsInt32(iRow, 1);
                Usuario.nombre = sl.GetCellValueAsString(iRow, 2);
                Usuario.edad = sl.GetCellValueAsInt32(iRow, 3);
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
            if (list2.Count() != 0)  
            {
                //se muestra los datos en el data grid view
                GridView1.DataSource = list2;
                GridView1.DataBind();
            }
            else 
            {
                //se muestra los datos en el data grid view
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
        public void Button1_Click(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile)
            {
                if (ChecarExtension(FileUpload1.FileName))
                {
                    FileUpload1.SaveAs(MapPath("~/ArchivoExcel/" + FileUpload1.FileName));
                    Label1.Text = FileUpload1.FileName + " archivo cargado exitosamente";
                    lblOculto.Text = MapPath("~/ArchivoExcel/" + FileUpload1.FileName);
                    try
                    {
                        CargarDatos(lblOculto.Text);
                    }
                    catch
                    {
                        Response.Write("Ocurrio un error debe cargar antes el archivo");
                    }
                }
            }
            else
            {
                Label1.Text = "Error al subir el archivo excel";
            }
        }
        //Colorea las filas que estan vacias al cargar el excel
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
                    Button3.Visible = false;
                    e.Row.BackColor = System.Drawing.Color.Red;
                    Label1.Text = "Llenar todos los campos vacios de la tabla por favor";
                }           
            }
        }
        protected void Button3_Click(object sender, EventArgs e)
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
                    var oPersona = new Persona();
                    oPersona.codigo = codigo;
                    oPersona.nombre = nombre;
                    oPersona.edad = edad;
                    db.Persona.Add(oPersona);
                    db.SaveChanges();
                    iRow++;
                    Label1.Text = "Datos Guardados Exitosamente";
                }
            }
        }
    }
}