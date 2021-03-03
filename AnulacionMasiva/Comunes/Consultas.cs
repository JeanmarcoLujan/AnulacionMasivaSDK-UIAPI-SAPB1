using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnulacionMasiva.Comunes
{
    class Consultas
    {

        #region Atributos
        private static StringBuilder m_sSQL = new StringBuilder(); //Variable para la construccion de strings
        #endregion


        public static string Facturas()
        {

            m_sSQL.Length = 0;
            m_sSQL.Append("	SELECT 'N' AS Seleccion  ,DocEntry, CardCode, CardName, DocTotal, DocDate, TaxDate, DocDueDate FROM OINV WHERE CANCELED = 'N' AND FolioNum is null and FolioNum is null ORDER BY DocDate DESC  ");

            return m_sSQL.ToString();
        }

    }
}
