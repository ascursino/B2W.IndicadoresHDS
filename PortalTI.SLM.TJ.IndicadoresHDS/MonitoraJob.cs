using System;
using System.Linq;
using System.Collections.Generic;
using System.Net.Security;
using System.ServiceModel.Channels;
using System.Net;
using System.Xml;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Microsoft.SharePoint.Administration;
using PortalTI.SLM.TJ.IndicadoresHDS.svcListDataHDS;

namespace PortalTI.SLM.TJ.IndicadoresHDS
{
    public class MonitoraJob : SPJobDefinition
    {
        //define listas e variáveis
        List<mdIndicadoresLinha> lstIndicadorLinhaResult = new List<mdIndicadoresLinha>();
        List<mdView> lstViewResult = new List<mdView>();
        mdConfig itemConfig = new mdConfig();

        public MonitoraJob() : base() { }

        public MonitoraJob(string jobName, SPService service, SPServer server, SPJobLockType targetType) :
            base(jobName, service, server, targetType) { }

        public MonitoraJob(string jobName, SPWebApplication webApplication)
            : base(jobName, webApplication, null, SPJobLockType.ContentDatabase)
        { this.Title = "B2W-Indicadores-HDS"; }

        public override void Execute(Guid contentDbId)
        {
            //Busca dados do HDS
            ExecuteSvc();

            //Gera indicadores
            GeraIndicador();

            //Verifica se tem alguma localidade faltando
            VerificaLocalidades();

            //Grava resultados na lista do Sharepoint
            GravaListaSharepoint();
        }


        public void GeraIndicador()
        {
            string varApelido = String.Empty;

            PortalSLMDataContext dc = new PortalSLMDataContext(new Uri(itemConfig.uri));
            dc.Credentials = new NetworkCredential(itemConfig.user, itemConfig.password, itemConfig.domain);            
            
            // Verifica a qtde de incidente e tarefa de catálogo (requisição), agrupa por localidade. 
            // Mostra incidente e tarefa de catálogo como registros separados.
            var queryView = (from lstView in lstViewResult
                             group lstView by new { lstView.Localidade } into lstGroup
                             select new
                             {
                                 Localidade = lstGroup.Key.Localidade,
                                 QtdeIncidente = lstGroup.Count(x => x.Tipo == "incident"),
                                 QtdeCatalogo = lstGroup.Count(x => x.Tipo == "sc_task")
                             });

            mdIndicadoresLinha itemIndicador = null;

            foreach (var itemQuery in queryView)
            {
                var queryLocal = (from lstLocal in dc.Localidade
                                  where lstLocal.Localidade.Equals(itemQuery.Localidade)
                                  select new { ApelidoLocal = lstLocal.Apelido });

                if (queryLocal.Count() == 0)
                {
                    varApelido = itemQuery.Localidade;
                }
                else
                {
                    foreach (var itemQueryLocal in queryLocal)
                    {
                        varApelido = itemQueryLocal.ApelidoLocal;
                    }
                }

                //insere incidente
                itemIndicador = new mdIndicadoresLinha();

                itemIndicador.NomeIndicador = "Chamados Ativos";
                itemIndicador.Localidade = itemQuery.Localidade;
                itemIndicador.NomeLocal = varApelido;
                itemIndicador.Texto = varApelido + " - Incidente";
                itemIndicador.Valor = itemQuery.QtdeIncidente.ToString();
                itemIndicador.Status = "UNDEFINED";
                itemIndicador.Data = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                lstIndicadorLinhaResult.Add(itemIndicador);

                //insere Requisição (tarefa de catálogo)
                itemIndicador = new mdIndicadoresLinha();

                itemIndicador.NomeIndicador = "Chamados Ativos";
                itemIndicador.Localidade = itemQuery.Localidade;
                itemIndicador.NomeLocal = varApelido;
                itemIndicador.Texto = varApelido + " - Requisição";
                itemIndicador.Valor = itemQuery.QtdeCatalogo.ToString();
                itemIndicador.Status = "UNDEFINED";
                itemIndicador.Data = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                lstIndicadorLinhaResult.Add(itemIndicador);
            }
        }

        public void GravaListaSharepoint()
        {
            PortalSLMDataContext dc = new PortalSLMDataContext(new Uri(itemConfig.uri));
            dc.Credentials = new NetworkCredential(itemConfig.user, itemConfig.password, itemConfig.domain);

            //deleta registros da lista
            foreach (IndicadorHDSItem deleteitem in dc.IndicadorHDS)
            {
                dc.DeleteObject(deleteitem);
                dc.SaveChanges();
            }

            //cria os novos registros gerados
            IndicadorHDSItem novoitem = null;

            foreach (mdIndicadoresLinha item in lstIndicadorLinhaResult)
            {
                novoitem = new IndicadorHDSItem();

                novoitem.NomeIndicador = item.NomeIndicador;
                novoitem.Titulo = item.NomeLocal;
                novoitem.Texto = item.Texto;
                novoitem.Valor = item.Valor;
                novoitem.Status = item.Status;
                novoitem.Data = item.Data;

                dc.AddToIndicadorHDS(novoitem);
                dc.SaveChanges();
            }

            lstIndicadorLinhaResult.Clear();
        }

        public void VerificaLocalidades()
        {
            PortalSLMDataContext dc = new PortalSLMDataContext(new Uri(itemConfig.uri));
            dc.Credentials = new NetworkCredential(itemConfig.user, itemConfig.password, itemConfig.domain);

            mdIndicadoresLinha itemIndicador = null;
            string varLocal = string.Empty;

            foreach (LocalidadeItem itemLista in dc.Localidade)
            {
                var queryView = (from lstIndicador in lstIndicadorLinhaResult
                    where lstIndicador.Localidade.Contains(itemLista.Localidade)
                    select lstIndicador.Localidade);

                if (queryView.Count() == 0)
                {
                    //insere incidente
                    itemIndicador = new mdIndicadoresLinha();

                    itemIndicador.NomeIndicador = "Chamados Ativos";
                    itemIndicador.Localidade = itemLista.Localidade;
                    itemIndicador.NomeLocal = itemLista.Apelido;
                    itemIndicador.Texto = itemLista.Apelido + " - Incidente";
                    itemIndicador.Valor = "0";
                    itemIndicador.Status = "UNDEFINED";
                    itemIndicador.Data = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                    lstIndicadorLinhaResult.Add(itemIndicador);

                    //insere tarefa de catálogo
                    itemIndicador = new mdIndicadoresLinha();

                    itemIndicador.NomeIndicador = "Chamados Ativos";
                    itemIndicador.Localidade = itemLista.Localidade;
                    itemIndicador.NomeLocal = itemLista.Apelido;
                    itemIndicador.Texto = itemLista.Apelido + " - Requisição";
                    itemIndicador.Valor = "0";
                    itemIndicador.Status = "UNDEFINED";
                    itemIndicador.Data = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                    lstIndicadorLinhaResult.Add(itemIndicador);
                }
            }
        }

        public void ExecuteSvc()
        {
            //------- SOMENTE PARA TESTE, POIS DESABILITA OS CERTIFICADOS DE SEGURANÇA ------------
            //ServicePointManager.ServerCertificateValidationCallback =
            //    delegate(object s, X509Certificate certificate,
            //             X509Chain chain, SslPolicyErrors sslPolicyErrors)
            //    { return true; };
            //------------------------------------------------------------

            mdView itemView = null;
            lstViewResult = new List<mdView>();

            HttpWebRequest request = CreateWebRequest();
            request.Credentials = new NetworkCredential("usr.kpi", "Fu13k0");

            XmlDocument soapEnvelopeXml = new XmlDocument();
            soapEnvelopeXml.LoadXml(@"
                <SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/""
                                   xmlns:u=""https://b2winc.service-now.com/u_vkpitask"">
                    <SOAP-ENV:Body>
                        <u:getRecords>
                            <__encoded_query>
                                tasktable_active=true^
                                tasktable_assignment_group=8fb17186611a41005fe69027bbb1f88f^
                                tasktable_stateNOT IN3,6^
                                tasktable_locationISNOTEMPTY^
                                tasktable_sys_class_name=incident^OR
                                tasktable_sys_class_name=sc_task
                            </__encoded_query>
                        </u:getRecords>
                    </SOAP-ENV:Body>
                </SOAP-ENV:Envelope>");

            using (Stream stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }

            using (WebResponse response = request.GetResponse())
            {
                using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                {
                    XmlReader reader = XmlReader.Create(rd);
                    string elementTemp = string.Empty;

                    while (reader.Read()) 
                    {
                        switch (reader.NodeType)
                        {
                            case XmlNodeType.Element:
                                switch (reader.Name)
                                {
                                    case "getRecordsResult":
                                        itemView = new mdView();
                                        break;
                                }
                                
                                elementTemp = reader.Name;
                                break;

                            case XmlNodeType.Text:
                                switch (elementTemp)
                                {
                                    case "grouptable_name":
                                        itemView.Grupo = reader.Value;
                                        break;

                                    case "locationtable_full_name":
                                        itemView.Localidade = reader.Value;
                                        break;

                                    case "tasktable_sys_class_name":
                                        itemView.Tipo = reader.Value;
                                        break;
                                }
                                break;

                            case XmlNodeType.EndElement:
                                switch (reader.Name)
                                {
                                    case "getRecordsResult":
                                        lstViewResult.Add(itemView);
                                        break;
                                }
                                break;
                        }
                    }
                }
            }
        }

        public HttpWebRequest CreateWebRequest()
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(@"https://b2winc.service-now.com/u_vkpitask.do?SOAP");
            webRequest.Headers.Add(@"SOAP:Action");
            webRequest.ContentType = "text/xml;charset=\"utf-8\"";
            webRequest.Accept = "text/xml";
            webRequest.Method = "POST";
            return webRequest;
        }
    }
}
