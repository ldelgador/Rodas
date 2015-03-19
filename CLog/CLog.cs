using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CLog
{
    public class CLog
    {
        private int _NivelLog=0;
        private string _FileLog = "DAL.log";
        

        public CLog(int pNivelLog, string pFileLog){
            _NivelLog = pNivelLog;
            _FileLog = pFileLog;
        }

        public  void TrataError(Exception pEx, String pSQL = "", int pNivelLog = 0, string pMensaje = "", bool pElevaError = true)
        {
            if (_NivelLog > pNivelLog )
            {
                //Excepcion ya registrada
                if (pEx.InnerException != null )
                {
                    if (pElevaError) throw pEx;
                }
                else
                {
                    RegistraError(pEx, pSQL, pNivelLog, pMensaje);                    
                    if (pElevaError) throw new Exception(pMensaje, pEx);                    
                }


            }
        }

        public void RegistraError(Exception pEx, String pSQL = "", int pNivelLog = 0, string pMensaje="")
        {
            if (_NivelLog > pNivelLog)
            {
                try
                {
                    using (StreamWriter w = File.AppendText(_FileLog))
                    {
                        w.WriteLine(" {0} {1}:", DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToLongTimeString(), w);
                        w.WriteLine("   Mensaje : {0}", pEx.Message);
                        if (_NivelLog >= 2)
                        {
                            w.WriteLine("   Pila : {0}", pEx.StackTrace);
                        }
                        if (pSQL != "" & _NivelLog >= 3)
                        {
                            w.WriteLine("   SQL : {0}", pSQL);
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("{0} {1} :", DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString());
                    Console.WriteLine(e.Message);
                    Console.WriteLine(e.StackTrace);
                    throw new Exception( "Error de registro de operaciones", e);
                }
            }            
        }

        private void HazLog(String pMensaje, StreamWriter pW)
        {            
            pW.WriteLine("   {0}", pMensaje);
        }
    }
}
