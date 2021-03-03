using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AnulacionMasiva.Comunes
{
    class Eventos_SBO
    {
        #region Constructores

        /// <summary>
        /// Constructor de la clase.
        /// </summary>
        public Eventos_SBO()
        {
            try
            {
                Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title = Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title.Replace(" #" + Conexion.Conexion_SBO.m_SBO_Appl.AppId.ToString(), "");
                Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title = Conexion.Conexion_SBO.m_SBO_Appl.Desktop.Title + " #" + Conexion.Conexion_SBO.m_SBO_Appl.AppId.ToString();
                RegistrarEventos();
               
                RegistrarMenu();
                Conexion.Conexion_SBO.m_SBO_Appl.StatusBar.SetText(AnulacionMasiva.Properties.Resources.NombreAddon + " Conectado con exito",
                    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }

        }

        #endregion

        #region Eventos

        void m_SBO_Appl_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                        RegistrarMenu();
                        break;
                }
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }
        void m_SBO_Appl_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
        void m_SBO_Appl_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.FormTypeEx)
                {
                    case "Frm_AMasivas":
                        Formularios.Frm_AMasivas oFrm_AMasivas = null;
                        oFrm_AMasivas = new AnulacionMasiva.Formularios.Frm_AMasivas(false);
                        oFrm_AMasivas.m_SBO_Appl_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        oFrm_AMasivas = null;
                        break;
                }
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }
      

        #endregion

        #region Metodos

        /// <summary>
        /// Método para registrar la opción del menú dentro del formulario de menus de SAP B1.
        /// </summary>
        private void RegistrarMenu()
        {
            try
            {

                CreaMenu("MSS_MPMD", "Anulaciones Masivas", "2048", SAPbouiCOM.BoMenuType.mt_POPUP); 
                CreaMenu("MSS_APMD", "Anulación Masiva de Facturas", "MSS_MPMD", SAPbouiCOM.BoMenuType.mt_STRING);//43538

            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        void m_SBO_Appl_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.BeforeAction)
                {
                    case false:
                        switch (pVal.MenuUID)
                        {
                            case "MSS_APMD":
                                Formularios.Frm_AMasivas oFrm_AMasivas = null;
                                oFrm_AMasivas = new AnulacionMasiva.Formularios.Frm_AMasivas(true);
                                oFrm_AMasivas = null;
                                break;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para registrar los eventos de la aplicación en SAP B1.
        /// </summary>
        private void RegistrarEventos()
        {
            try
            {
                Conexion.Conexion_SBO.m_SBO_Appl.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(m_SBO_Appl_AppEvent);
                Conexion.Conexion_SBO.m_SBO_Appl.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(m_SBO_Appl_FormDataEvent);
                Conexion.Conexion_SBO.m_SBO_Appl.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(m_SBO_Appl_ItemEvent);
                Conexion.Conexion_SBO.m_SBO_Appl.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(m_SBO_Appl_MenuEvent);
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        /// <summary>
        /// Método para registrar opciones de menú en el menu principal de SAP B1.
        /// </summary>
        /// <param name="uniqueId"></param>
        /// <param name="name"></param>
        /// <param name="principalMenuId"></param>
        /// <param name="type"></param>
        private void CreaMenu(string uniqueId, string name, string principalMenuId, SAPbouiCOM.BoMenuType type)
        {
            SAPbouiCOM.MenuCreationParams objParams;
            SAPbouiCOM.Menus objSubMenu;

            try
            {
                objSubMenu = Conexion.Conexion_SBO.m_SBO_Appl.Menus.Item(principalMenuId).SubMenus;

                if (Conexion.Conexion_SBO.m_SBO_Appl.Menus.Exists(uniqueId) == false)
                {
                    objParams = (SAPbouiCOM.MenuCreationParams)Conexion.Conexion_SBO.m_SBO_Appl.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objParams.Type = type;
                    objParams.UniqueID = uniqueId;
                    objParams.String = name;
                    objParams.Position = -1;
                    objSubMenu.AddEx(objParams);
                }
            }
            catch (Exception ex)
            {
                FuncionesComunes.DisplayErrorMessages(ex.Message, System.Reflection.MethodBase.GetCurrentMethod());
            }
        }

        #endregion
    }
}
