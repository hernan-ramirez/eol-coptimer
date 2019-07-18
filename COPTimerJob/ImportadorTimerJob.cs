using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;
using System.IO;
using System.Configuration;
using Microsoft.SharePoint;
using COPTimerJob.Importador;

namespace COPTimerJob
{
    class ImportadorTimerJob : SPJobDefinition
    {
        public ImportadorTimerJob()

            : base()
        {

        }

        public ImportadorTimerJob(string jobName, SPService service, SPServer server, SPJobLockType targetType)

            : base(jobName, service, server, targetType)
        {

        }

        public ImportadorTimerJob(string jobName, SPWebApplication webApplication)

            : base(jobName, webApplication, null, SPJobLockType.ContentDatabase)
        {

            this.Title = "Sincronizador Consultas Profesionales";

        }

        public override void Execute(Guid contentDbId)
        {
            ConnectionStringSettings connectionStringCodigos = System.Configuration.ConfigurationManager.ConnectionStrings["csSql"];
            AppSettingsReader app = new AppSettingsReader();
            string fecha = string.Concat(DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString("00"), DateTime.Now.Day.ToString("00"), DateTime.Now.Hour.ToString("00"), DateTime.Now.Minute.ToString("00"), DateTime.Now.Second.ToString("00"), (DateTime.Now.Millisecond).ToString("000"), "\\");
            string path = Path.Combine(@"F:\COP\GeneradosSql\", fecha);            
            //string path = Path.Combine(@"\\srv-kurma\COP\GeneradosSQL\", fecha); En el NOC está implementado éste path
            Directory.CreateDirectory(path);
            GeneradorDocumentos generador = new GeneradorDocumentos(
                "csSql",
                connectionStringCodigos.ConnectionString,
                path, true);
            generador.GenerarDocumentosDesdeDB();
            ImportadorFisicoMeta importador = new ImportadorFisicoMeta(path, app.GetValue("SiteUrl", typeof(string)).ToString(), "csSql", connectionStringCodigos.ConnectionString);
            importador.ImportarDocx();
        }

    }
}
