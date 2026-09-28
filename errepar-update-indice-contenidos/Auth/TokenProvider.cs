using System;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace Errepar.MetadataManager.Process.Auth
{
    public static class TokenProvider
    {
        // Método existente con client secret
        public static async Task<string> GetSharePointTokenAsync(string tenantId, string clientId, string clientSecret, string siteUrl)
        {
            var authority = $"https://login.microsoftonline.com/{tenantId}/";
            var app = ConfidentialClientApplicationBuilder
                        .Create(clientId)
                        .WithClientSecret(clientSecret)
                        .WithAuthority(authority)
                        .Build();

            var host = new Uri(siteUrl).Host;
            var scopes = new[] { $"https://{host}/.default" };

            var result = await app.AcquireTokenForClient(scopes).ExecuteAsync().ConfigureAwait(false);
            return result.AccessToken;
        }

        // Nuevo método con certificado (desde archivo .pfx)
        public static async Task<string> GetSharePointTokenWithCertificateAsync(string tenantId, string clientId, string certificatePath, string certificatePassword, string siteUrl)
        {
            var authority = $"https://login.microsoftonline.com/{tenantId}/";
            
            // Cargar certificado desde archivo .pfx
            var certificate = new X509Certificate2(certificatePath, certificatePassword, X509KeyStorageFlags.MachineKeySet);

            var app = ConfidentialClientApplicationBuilder
                        .Create(clientId)
                        .WithCertificate(certificate)
                        .WithAuthority(authority)
                        .Build();

            var host = new Uri(siteUrl).Host;
            var scopes = new[] { $"https://{host}/.default" };

            var result = await app.AcquireTokenForClient(scopes).ExecuteAsync().ConfigureAwait(false);
            return result.AccessToken;
        }

        // Método con certificado desde el almacén de Windows (Certificate Store)
        public static async Task<string> GetSharePointTokenWithCertificateThumbprintAsync(string tenantId, string clientId, string certificateThumbprint, string siteUrl)
        {
            var authority = $"https://login.microsoftonline.com/{tenantId}/";
            
            // Buscar certificado en el almacén de certificados de Windows por thumbprint
            X509Certificate2 certificate = null;
            using (var store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                var certs = store.Certificates.Find(X509FindType.FindByThumbprint, certificateThumbprint, false);
                
                if (certs.Count > 0)
                    certificate = certs[0];
                else
                    throw new InvalidOperationException($"Certificado con thumbprint '{certificateThumbprint}' no encontrado en el almacén.");
            }

            var app = ConfidentialClientApplicationBuilder
                        .Create(clientId)
                        .WithCertificate(certificate)
                        .WithAuthority(authority)
                        .Build();

            var host = new Uri(siteUrl).Host;
            var scopes = new[] { $"https://{host}/.default" };

            var result = await app.AcquireTokenForClient(scopes).ExecuteAsync().ConfigureAwait(false);
            return result.AccessToken;
        }
    }
}