using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DALFact = DALFactory.DALFactory;
using DAL;
using ET;
using System.Configuration;

namespace BL
{
    public class RodasBL
    {
        private string _CadenaConexion;
        private int _NivelLog;
        private string _FileLog;
        private DALFact.TipoFactory _TipoBD;
        private DAL.IRodasDAL _RodasDAL;
        private CLog.CLog _Log;
        private string _msjError ="Error de Lógica de Negocio.";

        public RodasBL()
        {
            _CadenaConexion = ConfigurationManager.ConnectionStrings["Rodas"].ConnectionString;
            _NivelLog = Convert.ToInt32(ConfigurationManager.AppSettings["NivelLog"]);
            _FileLog = ConfigurationManager.AppSettings["FileLog"];
            _TipoBD = (DALFact.TipoFactory)Convert.ToInt32(ConfigurationManager.AppSettings["TipoBD"]);

            _RodasDAL = DALFact.GetFactory(_TipoBD).getRodasDAL(_CadenaConexion, _NivelLog, _FileLog);
            _Log = new CLog.CLog(_NivelLog, _FileLog);
        }

        public bool CreaUsuario(Usuario pUsuario)
        {
            try
            {                                
                return _RodasDAL.CreaUsuario(pUsuario);
            }
            catch (Exception e)
            {
                _Log.TrataError(e,pMensaje:_msjError);
            }
            return false;
            
        }

    }
}
