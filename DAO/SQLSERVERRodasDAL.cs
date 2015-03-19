using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data.SqlClient;
using ET;

namespace DAL
{
    public class SQLSERVERRodasDAL: BaseDAL, IRodasDAL
    {
        
        public SQLSERVERRodasDAL(String pCadenaConexion, int pNivelLog=0, string pFileLog = "SQLSERVERRodasDAL.log")  :
            base(pCadenaConexion,pNivelLog,pFileLog)
        {
            
        }

        public bool CreaUsuario(Usuario pUsuario)
        {
            string sql = "INSERT INTO Usuarios(Nombre, Apellidos) VALUES (@p0, @p1)";
            try
            {
                using (var con = new SqlConnection(_CadenaConexion))
                {
                    con.Open();
                    
                    var query = new SqlCommand(sql, con);

                    query.Parameters.AddWithValue("@p0", pUsuario.Nombre);
                    query.Parameters.AddWithValue("@p1", pUsuario.Apellido);

                    query.ExecuteNonQuery();

                    return true;
                }
            }
            catch (Exception e)
            {
                _Log.TrataError(e,sql,pMensaje:_msjLog);                
            }
            return false;
        }

    }
}
