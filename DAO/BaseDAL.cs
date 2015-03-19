using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;


namespace DAL
{

    public class BaseDAL
    {
        protected string _CadenaConexion;       
        protected CLog.CLog _Log; 
        protected string _msjLog = "Error de Acceso a Datos";
        public BaseDAL() { }
        
        public BaseDAL(string pCadenaConexion, int pNivelLog, string pFileLog)
        {
            _CadenaConexion = pCadenaConexion;            
            _Log = new CLog.CLog(pNivelLog, pFileLog);

        }
       

       
    }
}
