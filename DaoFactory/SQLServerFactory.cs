using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL;


namespace DALFactory
{
    class SQLServerFactory: DALFactory
    {
        public override DAL.IRodasDAL getRodasDAL(string CadConexion, int NivelLog=0, string FileLog ="SQLSERVERRodasDAL.log")
        {
            //throw new NotImplementedException();
            return new DAL.SQLSERVERRodasDAL(CadConexion, NivelLog, FileLog);            
        }
    }
}
