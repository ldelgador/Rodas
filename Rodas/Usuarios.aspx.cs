using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using RodasBL = BL.RodasBL;
using ET;


namespace Rodas
{
    public partial class Usuarios : System.Web.UI.Page
    {
        
        protected void Page_Load(object sender, EventArgs e)
        {
            this.Label1.Visible = false;
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            try
            {
               
                RodasBL rodasBL = new RodasBL();
                Usuario usuario = new Usuario();

                usuario.Nombre = "Fulanito";
                usuario.Apellido = "De Copas";

                if (rodasBL.CreaUsuario(usuario))
                {
                    this.Label1.Text = "Usuario Creado";
                    this.Label1.Visible = true;
                }
                
            }
            catch (Exception ex )
            {
                Master.Log.TrataError(ex, pElevaError: false);
                this.Label1.Visible = true;                                                                                                
                this.Label1.Text = ex.Message;
            }
        }
    }
}