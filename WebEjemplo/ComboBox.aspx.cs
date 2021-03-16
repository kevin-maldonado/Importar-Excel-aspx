using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;

namespace WebEjemplo
{
    public partial class ComboBox : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e) 
        {
            
            
        }
        protected void Submit(object sender, EventArgs e)
        {
            string message = "";
            foreach (ListItem item in lstServicios.Items)
            {
                if (item.Selected)
                {
                    message += item.Text + " " + item.Value + "\\n";
                }
            }
            ClientScript.RegisterClientScriptBlock(this.GetType(), "alert", "alert('" + message + "');", true);
        }
    }
}