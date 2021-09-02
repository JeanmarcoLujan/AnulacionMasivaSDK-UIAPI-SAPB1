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


        public static string Facturas(string dt_FCDesde, string dt_FCHasta)
        {
            string s_Date = "";
            m_sSQL.Length = 0;
            m_sSQL.Append("SELECT 'N' AS Seleccion  ,DocEntry, CardCode, CardName, DocTotal, DocDate, TaxDate, DocDueDate FROM OINV WHERE CANCELED = 'N'  AND FolioNum is not null and FolioNum is not null AND DocStatus='O'  ");

            if (!dt_FCDesde.Equals(""))
            {
                if (!dt_FCHasta.Equals("")) { s_Date = "AND DocDate BETWEEN '" + dt_FCDesde + "' AND '" + dt_FCHasta + "'"; }
                else { s_Date = "AND DocDate >= '" + dt_FCDesde + "'"; }
            }
            else
            {
                if (!dt_FCHasta.Equals("")) { s_Date = "AND DocDate <= '" + dt_FCHasta + "'"; }
            }
            m_sSQL.AppendFormat("{0}  ORDER BY DocEntry ,DocDate DESC  ", s_Date);
            return m_sSQL.ToString();
        }

        public static string FormatoFechaC()
        {
            m_sSQL.Length = 0;
            m_sSQL.AppendFormat("SELECT [DateFormat] FROM OADM ");

            return m_sSQL.ToString();
        }

    }
}
