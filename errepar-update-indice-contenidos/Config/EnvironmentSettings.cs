using System;

namespace Errepar.MetadataManager.Process.Config
{
    public enum Ambiente
    {
        UAT,
        PROD
    }

    public class SharePointSettings
    {
        public string SiteUrl { get; init; } = string.Empty;
        public string TenantId { get; init; } = string.Empty;
        public string ClientId { get; init; } = string.Empty;
        public string CertificateThumbprint { get; init; } = string.Empty;
    }

    public class SolrSettings
    {
        public string UpdateUrl { get; init; } = string.Empty;
        public string User { get; init; } = string.Empty;
        public string Password { get; init; } = string.Empty;
    }

    public class MilvusSettings
    {
        public string AuthEndpoint { get; init; } = string.Empty;
        public string ApiKey { get; init; } = string.Empty;
        public string LoadAssetEndpoint { get; init; } = string.Empty;
        public string DeleteAssetEndpoint { get; init; } = string.Empty;
        public string SearchByTimestampEndpoint { get; init; } = string.Empty;
    }

    public class EnvironmentSettings
    {
        public Ambiente Ambiente { get; init; }
        public SharePointSettings SharePoint { get; init; } = new();
        public SolrSettings Solr { get; init; } = new();
        public MilvusSettings Milvus { get; init; } = new();
    }

    /// <summary>
    /// Configuración centralizada de cadenas de conexión, URLs, usuarios y credenciales
    /// por ambiente (UAT / PROD). Por defecto se utiliza UAT salvo que se indique
    /// explícitamente PROD como parámetro del proceso (argumento "--env=PROD" o
    /// variable de entorno METADATA_MANAGER_AMBIENTE=PROD).
    ///
    /// NOTA: los valores de PROD marcados con cadena vacía ("") no se encontraron de forma
    /// inequívoca en el código fuente original (estaban comentados, ausentes o con
    /// valores contradictorios). Deben completarse/confirmarse manualmente antes de usar PROD.
    /// </summary>
    public static class EnvironmentConfig
    {
        public static readonly EnvironmentSettings Uat = new()
        {
            Ambiente = Ambiente.UAT,
            SharePoint = new SharePointSettings
            {
                SiteUrl = "https://erreparsa.sharepoint.com/sites/ErreparDesarrollo",
                TenantId = "00f26ad1-2073-4746-a79f-c83061db35c0",
                ClientId = "8688eed4-7464-4288-9820-34849fd19296", // erreparDev
                CertificateThumbprint = "7BA783C08AEA10B3C35979E64388A404D4292ECA"
            },
            Solr = new SolrSettings
            {
                UpdateUrl = "https://solr.uat.errepar.com/solr/prodActivos02/update?commit=true",
                User = "uat-solr-acess",
                Password = "SXtuCde9&JW%pkAK"
            },
            Milvus = new MilvusSettings
            {
                AuthEndpoint = "https://accounts.uat.errepar.com/syserrepar/integration/authenticationandauthorization/e-auth-api/auth/job/getJobCredentialsEAuth",
                ApiKey = "e5169fe4-1c39-4356-b148-1f210cc2431b",
                LoadAssetEndpoint = "https://api.uat.errepar.com/syserrepar/embeddeddocumentai/loadAssetToMilvus",
                DeleteAssetEndpoint = "https://api.uat.errepar.com/syserrepar/embeddeddocumentai/deleteAssetMilvus",
                SearchByTimestampEndpoint = "https://api.uat.errepar.com/syserrepar/embeddeddocumentai/searchByTimestamp"
            }
        };

        public static readonly EnvironmentSettings Prod = new()
        {
            Ambiente = Ambiente.PROD,
           SharePoint = new SharePointSettings
            {
                // Fuente: DOCUMENTACION_METADATA_MANAGER.md ("Configuración de Ambientes" > "Producción")
                SiteUrl = "https://erreparsa.sharepoint.com/sites/Errepar",
                TenantId = "00f26ad1-2073-4746-a79f-c83061db35c0",
                ClientId = "8688eed4-7464-4288-9820-34849fd19296", // erreparDesarrollo (aparecía comentado en Program.cs)
                // No se encontró un thumbprint de certificado de PROD en el código ni en la documentación.
                CertificateThumbprint = "7BA783C08AEA10B3C35979E64388A404D4292ECA"
            },
            Solr = new SolrSettings
            {
                // Fuente: DOCUMENTACION_METADATA_MANAGER.md.
                // NOTA: en SolrManager.EliminarDeSolr aparecía además el host alternativo
                // "https://solr-prod-hcs.errepar.com/solr/prodActivos02/update?commit=true"; confirmar cuál está vigente.
                UpdateUrl = "https://solr.errepar.com/solr/prodActivos02/update?commit=true",
                //User = "admin",
                // NOTA: se encontraron dos passwords distintas para "admin" en el código
                // (SolrManager.EnviarASolr: "6s2HUXFb8la", SolrManager.EliminarDeSolr: "T0m4t1t02023*").
                // Se usa la documentada en DOCUMENTACION_METADATA_MANAGER.md; confirmar cuál es la vigente.
                //Password = "6s2HUXFb8la"
            
                User = "sa-sharepoint-prod",
                Password = "37jtgn%ds5RGJz$9pX7H7#mot"
            },
            Milvus = new MilvusSettings
            {
                // Encontrado comentado en MilvusManager.cs
                AuthEndpoint = "https://accounts.errepar.com/eauth/auth/job/getJobCredentialsEAuth",
                // No se encontró una API key distinta para PROD.
                ApiKey = "e5169fe4-1c39-4356-b148-1f210cc2431b",
                LoadAssetEndpoint = "https://api.errepar.com/syserrepar/embeddeddocumentai/loadAssetToMilvus",
                DeleteAssetEndpoint = "https://api.errepar.com/syserrepar/embeddeddocumentai/deleteAssetMilvus",
                SearchByTimestampEndpoint = "https://api.errepar.com/syserrepar/embeddeddocumentai/searchByTimestamp"
            }
        };

        /// <summary>
        /// Resuelve la configuración de ambiente a partir de un valor de texto ("UAT"/"PROD").
        /// Si el valor es nulo, vacío o no reconocido, se utiliza UAT por defecto.
        /// </summary>
        public static EnvironmentSettings Resolve(string? ambiente)
        {
            if (string.Equals(ambiente, "PROD", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(ambiente, "PRODUCCION", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(ambiente, "PRODUCTION", StringComparison.OrdinalIgnoreCase))
            {
                return Prod;
            }

            return Uat;
        }

        /// <summary>
        /// Resuelve el ambiente a partir de los argumentos del proceso ("--env=PROD", "--ambiente=PROD",
        /// o "PROD"/"UAT" como argumento suelto) y, si no se indica, de la variable de entorno
        /// METADATA_MANAGER_AMBIENTE. Por defecto se usa UAT.
        /// </summary>
        public static EnvironmentSettings ResolveFromArgs(string[]? args)
        {
            string? ambiente = "UAT";

            if (args != null)
            {
                foreach (var arg in args)
                {
                    if (arg.StartsWith("--env=", StringComparison.OrdinalIgnoreCase))
                        ambiente = arg["--env=".Length..];
                    else if (arg.StartsWith("--ambiente=", StringComparison.OrdinalIgnoreCase))
                        ambiente = arg["--ambiente=".Length..];
                    else if (string.Equals(arg, "PROD", StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(arg, "UAT", StringComparison.OrdinalIgnoreCase))
                        ambiente = arg;
                }
            }

            ambiente ??= Environment.GetEnvironmentVariable("METADATA_MANAGER_AMBIENTE");

            return Resolve(ambiente);
        }
    }
}
