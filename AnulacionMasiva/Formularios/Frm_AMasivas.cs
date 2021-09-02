using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnulacionMasiva.Formularios
{
    class Frm_AMasivas
    {
        public Frm_AMasivas(bool PrimeraCarga)
        {
            if (PrimeraCarga)
            {
                CargarFormulario();
                //CargarGridDocumentos();
            }
        }


        public void m_SBO_Appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.BeforeAction)
                {
                    case true:
                        switch (pVal.EventType)
                        {
                            
                        }
                        break;
                    case false:
                        switch (pVal.EventType)
                        {
                            case SAPbouiCOM.BoEventTypes.et_CLICK:
                                switch (pVal.ItemUID)
                                {
                                    case "grid_Res":
                                        if (pVal.ColUID.Equals("Seleccion"))
                                            MarcarHijos(pVal.Row, "grid_Res", "Seleccion");
                                        break;
                                    case "bt_anu": //bt_bu
                                        Form oForm = Conexion.Conexion_SBO.m_SBO_Appl.Forms.Item("Frm_AMasivas");

                                        AnularDocumentos(oForm);
                                        break;
                                    case "bt_bu": //bt_bu
                                        CargarGridDocumentos();
                                        break;
                                }
                                break;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        private void CargarFormulario()
        {
            SAPbouiCOM.Form oForm = null;//Objeto para manejar el formulario que vamos a crear.
            SAPbouiCOM.FormCreationParams oParams = null;//Objeto encargado de levantar el formulario almacenado en el archivo Accesorio.srf
            SAPbouiCOM.Item oItem = null; //variable para gestionar los objetos del formulario
            try
            {
                oParams = (SAPbouiCOM.FormCreationParams)Conexion.Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);//Generamos una instancia de tipo formulario
                oParams.FormType = "Frm_AMasivas";//Asignamos el tipo al formulario que deseamos levantar
                oParams.UniqueID = "Frm_AMasivas";//Creamos un ID unico para el formulario
                oParams.XmlData = AnulacionMasiva.Properties.Resources.Frm_AMasivas;//Cargamos el formulario del archivo Contador.srf contenido en nuestros recursos
                oForm = Conexion.Conexion_SBO.m_SBO_Appl.Forms.AddEx(oParams);//Registramos el formulario en SBO
                //oForm.DataSources.UserDataSources.Add("ud_FD", SAPbouiCOM.BoDataType.dt_DATE, 1);
                //oForm.DataSources.UserDataSources.Add("ud_FH", SAPbouiCOM.BoDataType.dt_DATE, 1);
                oForm.DataSources.UserDataSources.Item("ud_FA").Value = DateTime.Now.ToString("yyyyMMdd");
                oForm.Freeze(true);//Congela Ventana

                oForm.Left = Conexion.Conexion_SBO.m_SBO_Appl.Forms.GetFormByTypeAndCount(169, 0).Width + 10;
                oForm.Freeze(false);//Descongela la pantalla
                oForm.Visible = true;

            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("[66000-11]"))
                {
                    Conexion.Conexion_SBO.m_SBO_Appl.Forms.Item("Frm_AMasivas").Select();
                    Conexion.Conexion_SBO.m_SBO_Appl.Forms.Item("Frm_AMasivas").Visible = true;
                }
                else
                    Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }


        private void CargarGridDocumentos()
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Grid oGrid = null;
            SAPbobsCOM.AdminInfo oADM = null;
            SAPbobsCOM.CompanyService oCS = null;
            SAPbouiCOM.ProgressBar oPB = null;

            string s_FCDesde = "";
            string s_FCHasta = "";

            string formatoFecha = FormatoFecha();


            //cambio
            string sErrMsg = "";
            //cambio

            try
            {
                oForm = Conexion.Conexion_SBO.m_SBO_Appl.Forms.Item("Frm_AMasivas");
                oCS = Conexion.Conexion_SBO.m_oCompany.GetCompanyService();
                Conexion.Conexion_SBO.m_oCompany.GetCompanyDate();
                oADM = oCS.GetAdminInfo();


                oForm.Freeze(true);
                oItem = oForm.Items.Item("grid_Res");
                oForm.PaneLevel = 1;
                var asds = oForm.DataSources.UserDataSources.Item("ud_FD").Value.ToString();
                oGrid = (SAPbouiCOM.Grid)oItem.Specific;
                oGrid.DataTable = oForm.DataSources.DataTables.Item("dt_Res");
                

              
                oPB = Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.CreateProgressBar("Ajustando caracteristicas del Grid", 27, true);
                oPB.Value = 1;

                #region Operaciones de SN y Fechas

                if (!oForm.DataSources.UserDataSources.Item("ud_FD").Value.ToString().Equals(""))
                {
                    s_FCDesde = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("ud_FD").Value.ToString().Trim(), formatoFecha, null).Year.ToString();
                    s_FCDesde = s_FCDesde + CompletarCeros(DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("ud_FD").Value.ToString().Trim(), formatoFecha, null).Month.ToString());
                    s_FCDesde = s_FCDesde + CompletarCeros(DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("ud_FD").Value.ToString().Trim(), formatoFecha, null).Day.ToString());

                }

                if (!oForm.DataSources.UserDataSources.Item("ud_FH").Value.ToString().Equals(""))
                {
                    s_FCHasta = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("ud_FH").Value.ToString().Trim(), formatoFecha, null).Year.ToString();
                    s_FCHasta = s_FCHasta + CompletarCeros(DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("ud_FH").Value.ToString().Trim(), formatoFecha, null).Month.ToString());
                    s_FCHasta = s_FCHasta + CompletarCeros(DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("ud_FH").Value.ToString().Trim(), formatoFecha, null).Day.ToString());
                }


                oGrid.DataTable.ExecuteQuery(Comunes.Consultas.Facturas(s_FCDesde, s_FCHasta));

                oGrid.Columns.Item("Seleccion").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                oGrid.Columns.Item("Seleccion").TitleObject.Caption = "Seleccion";

                oGrid.Columns.Item("DocEntry").Editable = false;
                oGrid.Columns.Item("DocEntry").TitleObject.Caption = "DocEntry";
                ((SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("DocEntry")).LinkedObjectType = "13";

                System.Threading.Thread.Sleep(1000);
                oPB.Stop();
                Comunes.FuncionesComunes.LiberarObjetoGenerico(oPB);
                #endregion

                oGrid.AutoResizeColumns();
                //oGrid.Rows.CollapseAll();
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                if (oPB != null)
                {
                    oPB.Stop();
                    Comunes.FuncionesComunes.LiberarObjetoGenerico(oPB);
                }
            }
        }

        private string CompletarCeros(string s_valor)
        {
            if (s_valor.Length.ToString().Equals("1"))
                return "0" + s_valor;
            else
                return s_valor;
        }

        public  string FormatoFecha()
        {
            SAPbobsCOM.Recordset oRS = null;
            string Formato = "";
            try
            {

                oRS = (SAPbobsCOM.Recordset)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery(Comunes.Consultas.FormatoFechaC());
                if (oRS.RecordCount > 0)
                {
                    string result = oRS.Fields.Item(0).Value.ToString();
                    if (result.Equals("0"))
                    {
                        Formato = "dd/MM/yy";
                    }
                    else if (result.Equals("1"))
                    {
                        Formato = "dd/MM/yyyy";
                    }

                }


            }
            catch (Exception ex)
            {
                
            }
            return Formato;
        }

        private void MarcarHijos(int row, string grid, string columna)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Grid oGrid = null;
            string Valor = "Y";

            try
            {
                oForm = Conexion.Conexion_SBO.m_SBO_Appl.Forms.Item("Frm_AMasivas");
                //oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grd_SNList").Specific;
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(grid).Specific;

                oForm.Freeze(true);


                if (!row.Equals(-1))
                {

                    for (int i = 0; i < oGrid.Rows.Count; i++)
                    {
                        if (oGrid.DataTable.GetValue(columna, oGrid.GetDataTableRowIndex(i)).ToString().Trim().Equals("Y"))
                        {
                            Valor = "N";
                            break;
                        }
                        Valor = "Y";
                    }
                }
                else
                {

                    var estado = oGrid.DataTable.GetValue(columna, oGrid.GetDataTableRowIndex(1)).ToString().Trim();
                    if (estado.Equals("Y"))
                        Valor = "N";
                    else
                        Valor = "Y";
                    for (int i = 0; i < oGrid.Rows.Count; i++)
                    {


                        oGrid.DataTable.SetValue(columna, oGrid.GetDataTableRowIndex(i), Valor);
                    }
                }

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }


        private void AnularDocumentos(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.StaticText oStatic = null;

            string sErrMsg = "";
            SAPbouiCOM.ProgressBar oPB = null;

            SAPbobsCOM.Recordset oRS = null;
    
            SAPbobsCOM.AdminInfo oADM = null;
            SAPbobsCOM.CompanyService oCS = null;

            SAPbobsCOM.Documents oDocument = null;
            SAPbobsCOM.Documents oNotaCredito = null;

            string formatoFecha = FormatoFecha();
            try
            {
                oForm.Freeze(true);
                oRS = (SAPbobsCOM.Recordset)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oCS = Conexion.Conexion_SBO.m_oCompany.GetCompanyService();
                oADM = oCS.GetAdminInfo();

                #region Ajustar DataTable

                // Limpiando documentos no seleccionados
                

                for (int j = 0; j < oForm.DataSources.DataTables.Item("dt_Res").Rows.Count; j++)
                {
                    if (!oForm.DataSources.DataTables.Item("dt_Res").GetValue("Seleccion", j).ToString().Equals("N"))
                    {

                        oDocument = null;
                        oDocument = (SAPbobsCOM.Documents)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                        oDocument.GetByKey(int.Parse(oForm.DataSources.DataTables.Item("dt_Res").GetValue("DocEntry", j).ToString()));
                        oNotaCredito = null;
                        oNotaCredito = (SAPbobsCOM.Documents)Conexion.Conexion_SBO.m_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);

                        oNotaCredito.CardCode = oDocument.CardCode;

                        oNotaCredito.DocDate = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("ud_FA").Value.ToString().Trim(), formatoFecha, null); 
                        oNotaCredito.TaxDate = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("ud_FA").Value.ToString().Trim(), formatoFecha, null);
                        oNotaCredito.DocDueDate = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("ud_FA").Value.ToString().Trim(), formatoFecha, null);
                        oNotaCredito.Indicator = "07";

                        //oNotaCredito.FolioPrefixString = "LT";
                        //oNotaCredito.FolioNumber = oDocument.FolioNumber;


                        for (int i=0; i < oNotaCredito.Lines.Count; i++) {
                            oNotaCredito.Lines.SetCurrentLine(i);
                            oNotaCredito.Lines.BaseEntry = oDocument.DocEntry;
                            oNotaCredito.Lines.BaseLine = oDocument.Lines.LineNum;
                            oNotaCredito.Lines.WarehouseCode = oDocument.Lines.WarehouseCode;
                            oNotaCredito.Lines.AccountCode = oDocument.Lines.AccountCode;
                            oNotaCredito.Lines.BaseType = 13;
                            
                            if (i != oNotaCredito.Lines.Count-1)
                                oNotaCredito.Lines.Add();

                        }
                        var iErrCod = oNotaCredito.Add();

                        if (iErrCod != 0)
                        {
                            Conexion.Conexion_SBO.m_oCompany.GetLastError(out iErrCod, out sErrMsg);

                        }
                        
                            
                        


                    }
                }

                CargarGridDocumentos();

                

                #endregion

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                Comunes.FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
                oForm.Freeze(false);

                if (oPB != null)
                {
                    oPB.Stop();
                    Comunes.FuncionesComunes.LiberarObjetoGenerico(oPB);
                }
            }
        }
    }
}
