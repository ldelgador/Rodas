using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL;
using DALFactory;


namespace DALFactory
{
    
    public  abstract class DALFactory
    {
        public enum TipoFactory
        { 
            SQLSERVER = 1,
            ORACLE = 2,            
            MYSQL = 3
        }
        public static DALFactory GetFactory(TipoFactory Tipo)
        {
            switch (Tipo)
            {
                case TipoFactory.ORACLE:                    
                    break;
                case TipoFactory.SQLSERVER:
                    return  new SQLServerFactory();
                    //break;
                case TipoFactory.MYSQL:
                    break;
                default:
                    break;
            }
            return null;
        }
         public abstract DAL.IRodasDAL getRodasDAL(string CadConexion, int NivelLog = 0, string FileLog="RodasDAL.log");

    } // DAOFactory


}
