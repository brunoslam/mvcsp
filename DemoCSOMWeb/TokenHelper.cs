using Microsoft.IdentityModel;
using Microsoft.IdentityModel.SecurityTokenService;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.IdentityModel.S2S.Tokens;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IdentityModel.Selectors;
using System.IdentityModel.Tokens;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Security.Principal;
using System.ServiceModel;
using System.Text;
using System.Web;
using System.Web.Configuration;
using System.Web.Script.Serialization;
using AudienceRestriction = Microsoft.IdentityModel.Tokens.AudienceRestriction;
using AudienceUriValidationFailedException = Microsoft.IdentityModel.Tokens.AudienceUriValidationFailedException;
using SecurityTokenHandlerConfiguration = Microsoft.IdentityModel.Tokens.SecurityTokenHandlerConfiguration;
using X509SigningCredentials = Microsoft.IdentityModel.SecurityTokenService.X509SigningCredentials;

namespace DemoCSOMWeb
{
    public static class TokenHelper
    {
        #region campos públicos

        /// <summary>
        /// Entidad de seguridad de SharePoint.
        /// </summary>
        public const string SharePointPrincipal = "00000003-0000-0ff1-ce00-000000000000";

        /// <summary>
        /// Duración del token de acceso HighTrust, 12 horas.
        /// </summary>
        public static readonly TimeSpan HighTrustAccessTokenLifetime = TimeSpan.FromHours(12.0);

        #endregion public fields

        #region métodos públicos

        /// <summary>
        /// Recupera la cadena de token de contexto de la solicitud especificada; para ello, busca los nombres de parámetro conocidos en los 
        /// parámetros del formulario expuestos (con POST) y en la cadena de consulta. Devuelve NULL si no se encuentra ningún token de contexto.
        /// </summary>
        /// <param name="request">HttpRequest donde se va buscar un token de contexto</param>
        /// <returns>Cadena de token de contexto</returns>
        public static string GetContextTokenFromRequest(HttpRequest request)
        {
            return GetContextTokenFromRequest(new HttpRequestWrapper(request));
        }

        /// <summary>
        /// Recupera la cadena de token de contexto de la solicitud especificada; para ello, busca los nombres de parámetro conocidos en los 
        /// parámetros del formulario expuestos (con POST) y en la cadena de consulta. Devuelve NULL si no se encuentra ningún token de contexto.
        /// </summary>
        /// <param name="request">HttpRequest donde se va buscar un token de contexto</param>
        /// <returns>Cadena de token de contexto</returns>
        public static string GetContextTokenFromRequest(HttpRequestBase request)
        {
            string[] paramNames = { "AppContext", "AppContextToken", "AccessToken", "SPAppToken" };
            foreach (string paramName in paramNames)
            {
                if (!string.IsNullOrEmpty(request.Form[paramName]))
                {
                    return request.Form[paramName];
                }
                if (!string.IsNullOrEmpty(request.QueryString[paramName]))
                {
                    return request.QueryString[paramName];
                }
            }
            return null;
        }

        /// <summary>
        /// Valide que una cadena de token de contexto especificada está pensada para esta aplicación según los parámetros 
        /// especificados en el archivo web.config. Entre los parámetros de web.config usados para la validación se incluyen ClientId, 
        /// HostedAppHostNameOverride, HostedAppHostName, ClientSecret y Realm (si se especifica). Si HostedAppHostNameOverride está presente,
        /// se usará para la validación. De lo contrario, si <paramref name="appHostName"/> no es 
        /// NULL, se usa para la validación en lugar del HostedAppHostName del archivo web.config. Si el token no es válido, se 
        /// producirá una excepción. Si el token es válido, se actualizará la dirección URL de los metadatos estáticos de STS de TokenHelper según el contenido del token
        /// y se devolverá un JsonWebSecurityToken basado en el token de contexto.
        /// </summary>
        /// <param name="contextTokenString">Token de contexto que se va a validar</param>
        /// <param name="appHostName">Autoridad de URL, que consta del nombre de host del Sistema de nombres de dominio (DNS) o la dirección IP y el número de puerto, que se va a usar para la validación de la audiencia del token.
        /// Si es NULL, se usará en su lugar la configuración de HostedAppHostName del archivo web.config. La configuración de HostedAppHostNameOverride del archivo web.config, si está presente, se usará 
        /// para la validación en lugar de <paramref name="appHostName"/> .</param>
        /// <returns>JsonWebSecurityToken basado en el token de contexto.</returns>
        public static SharePointContextToken ReadAndValidateContextToken(string contextTokenString, string appHostName = null)
        {
            JsonWebSecurityTokenHandler tokenHandler = CreateJsonWebSecurityTokenHandler();
            SecurityToken securityToken = tokenHandler.ReadToken(contextTokenString);
            JsonWebSecurityToken jsonToken = securityToken as JsonWebSecurityToken;
            SharePointContextToken token = SharePointContextToken.Create(jsonToken);

            string stsAuthority = (new Uri(token.SecurityTokenServiceUri)).Authority;
            int firstDot = stsAuthority.IndexOf('.');

            GlobalEndPointPrefix = stsAuthority.Substring(0, firstDot);
            AcsHostUrl = stsAuthority.Substring(firstDot + 1);

            tokenHandler.ValidateToken(jsonToken);

            string[] acceptableAudiences;
            if (!String.IsNullOrEmpty(HostedAppHostNameOverride))
            {
                acceptableAudiences = HostedAppHostNameOverride.Split(';');
            }
            else if (appHostName == null)
            {
                acceptableAudiences = new[] { HostedAppHostName };
            }
            else
            {
                acceptableAudiences = new[] { appHostName };
            }

            bool validationSuccessful = false;
            string realm = Realm ?? token.Realm;
            foreach (var audience in acceptableAudiences)
            {
                string principal = GetFormattedPrincipal(ClientId, audience, realm);
                if (StringComparer.OrdinalIgnoreCase.Equals(token.Audience, principal))
                {
                    validationSuccessful = true;
                    break;
                }
            }

            if (!validationSuccessful)
            {
                throw new AudienceUriValidationFailedException(
                    String.Format(CultureInfo.CurrentCulture,
                    "\"{0}\" is not the intended audience \"{1}\"", String.Join(";", acceptableAudiences), token.Audience));
            }

            return token;
        }

        /// <summary>
        /// Recupera un token de acceso de ACS para llamar al origen del token de contexto especificado en el 
        /// targetHost indicado. targetHost debe estar registrado para la entidad de seguridad que envió el token de contexto.
        /// </summary>
        /// <param name="contextToken">Token de contexto emitido por la audiencia de token de acceso prevista</param>
        /// <param name="targetHost">Autoridad de URL de la entidad de seguridad de destino</param>
        /// <returns>Token de acceso cuya audiencia coincide con el origen del token de contexto</returns>
        public static OAuth2AccessTokenResponse GetAccessToken(SharePointContextToken contextToken, string targetHost)
        {
            string targetPrincipalName = contextToken.TargetPrincipalName;

            // Extrae el refreshToken del token de contexto
            string refreshToken = contextToken.RefreshToken;

            if (String.IsNullOrEmpty(refreshToken))
            {
                return null;
            }

            string targetRealm = Realm ?? contextToken.Realm;

            return GetAccessToken(refreshToken,
                                  targetPrincipalName,
                                  targetHost,
                                  targetRealm);
        }

        /// <summary>
        /// Usa el código de autorización especificado para recuperar un token de acceso de ACS para llamar a la entidad de seguridad indicada 
        /// en el targetHost especificado. targetHost debe estar registrado para la entidad de seguridad de destino.  Si el dominio kerberos es 
        /// NULL, se usará en su lugar la configuración de "Realm" del archivo web.config.
        /// </summary>
        /// <param name="authorizationCode">Código de autorización para intercambiar el código de acceso</param>
        /// <param name="targetPrincipalName">Nombre de la entidad de seguridad de destino para la que se va a recuperar un token de acceso</param>
        /// <param name="targetHost">Autoridad de URL de la entidad de seguridad de destino</param>
        /// <param name="targetRealm">Dominio kerberos que se va a usar para el identificador de nombre y la audiencia del token de acceso</param>
        /// <param name="redirectUri">URI de redirección registrada para este complemento</param>
        /// <returns>Token de acceso con una audiencia de la entidad de seguridad de destino</returns>
        public static OAuth2AccessTokenResponse GetAccessToken(
            string authorizationCode,
            string targetPrincipalName,
            string targetHost,
            string targetRealm,
            Uri redirectUri)
        {
            if (targetRealm == null)
            {
                targetRealm = Realm;
            }

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = GetFormattedPrincipal(ClientId, null, targetRealm);

            // Crear solicitud para el token. RedirectUri es NULL.  Se producirá un error si el URI de redirección está registrado
            OAuth2AccessTokenRequest oauth2Request =
                OAuth2MessageFactory.CreateAccessTokenRequestWithAuthorizationCode(
                    clientId,
                    ClientSecret,
                    authorizationCode,
                    redirectUri,
                    resource);

            // Obtener token
            OAuth2S2SClient client = new OAuth2S2SClient();
            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (RequestFailedException)
            {
                if (!string.IsNullOrEmpty(SecondaryClientSecret))
                {
                    oauth2Request =
                    OAuth2MessageFactory.CreateAccessTokenRequestWithAuthorizationCode(
                        clientId,
                        SecondaryClientSecret,
                        authorizationCode,
                        redirectUri,
                        resource);

                    oauth2Response =
                        client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
                }
                else
                {
                    throw;
                }
            }
            catch (WebException wex)
            {
                using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }

            return oauth2Response;
        }

        /// <summary>
        /// Usa el token de actualización especificado con el fin de recuperar un token de acceso de ACS para llamar a la entidad de seguridad especificada 
        /// en el targetHost especificado. targetHost debe estar registrado para la entidad de seguridad de destino.  Si el dominio kerberos es 
        /// NULL, se usará en su lugar la configuración de "Realm" del archivo web.config.
        /// </summary>
        /// <param name="refreshToken">Token de actualización para intercambiar el token de acceso</param>
        /// <param name="targetPrincipalName">Nombre de la entidad de seguridad de destino para la que se va a recuperar un token de acceso</param>
        /// <param name="targetHost">Autoridad de URL de la entidad de seguridad de destino</param>
        /// <param name="targetRealm">Dominio kerberos que se va a usar para el identificador de nombre y la audiencia del token de acceso</param>
        /// <returns>Token de acceso con una audiencia de la entidad de seguridad de destino</returns>
        public static OAuth2AccessTokenResponse GetAccessToken(
            string refreshToken,
            string targetPrincipalName,
            string targetHost,
            string targetRealm)
        {
            if (targetRealm == null)
            {
                targetRealm = Realm;
            }

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = GetFormattedPrincipal(ClientId, null, targetRealm);

            OAuth2AccessTokenRequest oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithRefreshToken(clientId, ClientSecret, refreshToken, resource);

            // Obtener token
            OAuth2S2SClient client = new OAuth2S2SClient();
            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (RequestFailedException)
            {
                if (!string.IsNullOrEmpty(SecondaryClientSecret))
                {
                    oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithRefreshToken(clientId, SecondaryClientSecret, refreshToken, resource);
                    oauth2Response =
                        client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
                }
                else
                {
                    throw;
                }
            }
            catch (WebException wex)
            {
                using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }

            return oauth2Response;
        }

        /// <summary>
        /// Recupera un token de acceso solo de aplicación de ACS para llamar a la entidad de seguridad especificada 
        /// en el targetHost especificado. targetHost debe estar registrado para la entidad de seguridad de destino.  Si el dominio kerberos es 
        /// NULL, se usará en su lugar la configuración de "Realm" del archivo web.config.
        /// </summary>
        /// <param name="targetPrincipalName">Nombre de la entidad de seguridad de destino para la que se va a recuperar un token de acceso</param>
        /// <param name="targetHost">Autoridad de URL de la entidad de seguridad de destino</param>
        /// <param name="targetRealm">Dominio kerberos que se va a usar para el identificador de nombre y la audiencia del token de acceso</param>
        /// <returns>Token de acceso con una audiencia de la entidad de seguridad de destino</returns>
        public static OAuth2AccessTokenResponse GetAppOnlyAccessToken(
            string targetPrincipalName,
            string targetHost,
            string targetRealm)
        {

            if (targetRealm == null)
            {
                targetRealm = Realm;
            }

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = GetFormattedPrincipal(ClientId, HostedAppHostName, targetRealm);

            OAuth2AccessTokenRequest oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithClientCredentials(clientId, ClientSecret, resource);
            oauth2Request.Resource = resource;

            // Obtener token
            OAuth2S2SClient client = new OAuth2S2SClient();

            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (RequestFailedException)
            {
                if (!string.IsNullOrEmpty(SecondaryClientSecret))
                {
                    oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithClientCredentials(clientId, SecondaryClientSecret, resource);
                    oauth2Request.Resource = resource;

                    oauth2Response =
                        client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
                }
                else
                {
                    throw;
                }
            }
            catch (WebException wex)
            {
                using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }

            return oauth2Response;
        }

        /// <summary>
        /// Crea un contexto de cliente basado en las propiedades de un receptor de eventos remotos
        /// </summary>
        /// <param name="properties">Propiedades de un receptor de eventos remotos</param>
        /// <returns>ClientContext preparado para llamar al sitio web donde se originó el evento</returns>
        public static ClientContext CreateRemoteEventReceiverClientContext(SPRemoteEventProperties properties)
        {
            Uri sharepointUrl;
            if (properties.ListEventProperties != null)
            {
                sharepointUrl = new Uri(properties.ListEventProperties.WebUrl);
            }
            else if (properties.ItemEventProperties != null)
            {
                sharepointUrl = new Uri(properties.ItemEventProperties.WebUrl);
            }
            else if (properties.WebEventProperties != null)
            {
                sharepointUrl = new Uri(properties.WebEventProperties.FullUrl);
            }
            else
            {
                return null;
            }

            if (IsHighTrustApp())
            {
                return GetS2SClientContextWithWindowsIdentity(sharepointUrl, null);
            }

            return CreateAcsClientContextForUrl(properties, sharepointUrl);
        }

        /// <summary>
        /// Crea un contexto de cliente basado en las propiedades de un evento de complemento
        /// </summary>
        /// <param name="properties">Propiedades de un evento de complemento</param>
        /// <param name="useAppWeb">Es True para que el destino sea el sitio web de aplicación, false para que sea el sitio web host</param>
        /// <returns>ClientContext preparado para llamar al sitio web de aplicación o al sitio web primario</returns>
        public static ClientContext CreateAppEventClientContext(SPRemoteEventProperties properties, bool useAppWeb)
        {
            if (properties.AppEventProperties == null)
            {
                return null;
            }

            Uri sharepointUrl = useAppWeb ? properties.AppEventProperties.AppWebFullUrl : properties.AppEventProperties.HostWebFullUrl;
            if (IsHighTrustApp())
            {
                return GetS2SClientContextWithWindowsIdentity(sharepointUrl, null);
            }

            return CreateAcsClientContextForUrl(properties, sharepointUrl);
        }

        /// <summary>
        /// Recupera un token de acceso de ACS usando el código de autorización especificado y usa dicho token para 
        /// crear un contexto de cliente
        /// </summary>
        /// <param name="targetUrl">Dirección URL del sitio de SharePoint de destino</param>
        /// <param name="authorizationCode">Código de autorización que se va a usar al recuperar el token de acceso de ACS</param>
        /// <param name="redirectUri">URI de redirección registrada para este complemento</param>
        /// <returns>ClientContext preparado para llamar a targetUrl con un token de acceso válido</returns>
        public static ClientContext GetClientContextWithAuthorizationCode(
            string targetUrl,
            string authorizationCode,
            Uri redirectUri)
        {
            return GetClientContextWithAuthorizationCode(targetUrl, SharePointPrincipal, authorizationCode, GetRealmFromTargetUrl(new Uri(targetUrl)), redirectUri);
        }

        /// <summary>
        /// Recupera un token de acceso de ACS usando el código de autorización especificado y usa dicho token para 
        /// crear un contexto de cliente
        /// </summary>
        /// <param name="targetUrl">Dirección URL del sitio de SharePoint de destino</param>
        /// <param name="targetPrincipalName">Nombre de la entidad de seguridad de SharePoint de destino</param>
        /// <param name="authorizationCode">Código de autorización que se va a usar al recuperar el token de acceso de ACS</param>
        /// <param name="targetRealm">Dominio kerberos que se va a usar para el identificador de nombre y la audiencia del token de acceso</param>
        /// <param name="redirectUri">URI de redirección registrada para este complemento</param>
        /// <returns>ClientContext preparado para llamar a targetUrl con un token de acceso válido</returns>
        public static ClientContext GetClientContextWithAuthorizationCode(
            string targetUrl,
            string targetPrincipalName,
            string authorizationCode,
            string targetRealm,
            Uri redirectUri)
        {
            Uri targetUri = new Uri(targetUrl);

            string accessToken =
                GetAccessToken(authorizationCode, targetPrincipalName, targetUri.Authority, targetRealm, redirectUri).AccessToken;

            return GetClientContextWithAccessToken(targetUrl, accessToken);
        }

        /// <summary>
        /// Usa el token de acceso especificado para crear un contexto de cliente
        /// </summary>
        /// <param name="targetUrl">Dirección URL del sitio de SharePoint de destino</param>
        /// <param name="accessToken">Token de acceso que se va a usar al llamar al targetUrl especificado</param>
        /// <returns>ClientContext preparado para llamar a targetUrl con el token de acceso especificado</returns>
        public static ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken)
        {
            ClientContext clientContext = new ClientContext(targetUrl);

            clientContext.AuthenticationMode = ClientAuthenticationMode.Anonymous;
            clientContext.FormDigestHandlingEnabled = false;
            clientContext.ExecutingWebRequest +=
                delegate(object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken;
                };

            return clientContext;
        }

        /// <summary>
        /// Recupera un token de acceso de ACS usando el token de contexto especificado y usa ese token de acceso para crear
        /// un contexto de cliente
        /// </summary>
        /// <param name="targetUrl">Dirección URL del sitio de SharePoint de destino</param>
        /// <param name="contextTokenString">Token de contexto recibido del sitio de SharePoint de destino</param>
        /// <param name="appHostUrl">Autoridad URL del complemento hospedado. Si es nulo, el valor es el de HostedAppHostName.
        /// valor de HostedAppHostName del archivo web.config</param>
        /// <returns>ClientContext preparado para llamar a targetUrl con un token de acceso válido</returns>
        public static ClientContext GetClientContextWithContextToken(
            string targetUrl,
            string contextTokenString,
            string appHostUrl)
        {
            SharePointContextToken contextToken = ReadAndValidateContextToken(contextTokenString, appHostUrl);

            Uri targetUri = new Uri(targetUrl);

            string accessToken = GetAccessToken(contextToken, targetUri.Authority).AccessToken;

            return GetClientContextWithAccessToken(targetUrl, accessToken);
        }

        /// <summary>
        /// Devuelve la dirección URL de SharePoint a la que el complemento debe redirigir el explorador para solicitar su consentimiento.
        /// y obtener un código de autorización.
        /// </summary>
        /// <param name="contextUrl">Dirección URL absoluta del sitio de SharePoint</param>
        /// <param name="scope">Permisos delimitados por espacios para solicitar al sitio de SharePoint en formato "abreviado" 
        /// (por ejemplo, "Web.Read Site.Write")</param>
        /// <returns>Dirección URL de la página de autorización OAuth del sitio de SharePoint</returns>
        public static string GetAuthorizationUrl(string contextUrl, string scope)
        {
            return string.Format(
                "{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code",
                EnsureTrailingSlash(contextUrl),
                AuthorizationPage,
                ClientId,
                scope);
        }

        /// <summary>
        /// Devuelve la dirección URL de SharePoint a la que el complemento debe redirigir el explorador para solicitar su consentimiento.
        /// y obtener un código de autorización.
        /// </summary>
        /// <param name="contextUrl">Dirección URL absoluta del sitio de SharePoint</param>
        /// <param name="scope">Permisos delimitados por espacios para solicitar al sitio de SharePoint en formato "abreviado"
        /// (por ejemplo, "Web.Read Site.Write")</param>
        /// <param name="redirectUri">URI al que SharePoint debe redirigir el explorador después del consentimiento 
        /// concedido</param>
        /// <returns>Dirección URL de la página de autorización OAuth del sitio de SharePoint</returns>
        public static string GetAuthorizationUrl(string contextUrl, string scope, string redirectUri)
        {
            return string.Format(
                "{0}{1}?IsDlg=1&client_id={2}&scope={3}&response_type=code&redirect_uri={4}",
                EnsureTrailingSlash(contextUrl),
                AuthorizationPage,
                ClientId,
                scope,
                redirectUri);
        }

        /// <summary>
        /// Devuelve la dirección URL de SharePoint a la que el complemento debe redirigir el explorador para solicitar un nuevo token de contexto.
        /// </summary>
        /// <param name="contextUrl">Dirección URL absoluta del sitio de SharePoint</param>
        /// <param name="redirectUri">URI al que SharePoint debe redirigir el explorador con un token de contexto</param>
        /// <returns>Dirección URL de la página de redirección del token de contexto del sitio de SharePoint</returns>
        public static string GetAppContextTokenRequestUrl(string contextUrl, string redirectUri)
        {
            return string.Format(
                "{0}{1}?client_id={2}&redirect_uri={3}",
                EnsureTrailingSlash(contextUrl),
                RedirectPage,
                ClientId,
                redirectUri);
        }

        /// <summary>
        /// Recupera un token de acceso S2S firmado por el certificado privado de la aplicación en nombre de 
        /// WindowsIdentity que se ha especificado y destinado a SharePoint en el targetApplicationUri. Si no se especifica ningún valor de Realm en 
        /// web.config, se emitirá un desafío de autenticación al targetApplicationUri para detectarlo.
        /// </summary>
        /// <param name="targetApplicationUri">Dirección URL del sitio de SharePoint de destino</param>
        /// <param name="identity">Identidad de Windows del usuario en cuyo nombre se va a crear el token de acceso</param>
        /// <returns>Token de acceso con una audiencia de la entidad de seguridad de destino</returns>
        public static string GetS2SAccessTokenWithWindowsIdentity(
            Uri targetApplicationUri,
            WindowsIdentity identity)
        {
            string realm = string.IsNullOrEmpty(Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : Realm;

            JsonWebTokenClaim[] claims = identity != null ? GetClaimsWithWindowsIdentity(identity) : null;

            return GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);
        }

        /// <summary>
        /// Recupera un contexto de cliente S2S con un token de acceso firmado por el certificado privado de la aplicación en 
        /// nombre de WindowsIdentity que se ha especificado y destinado a la aplicación en el targetApplicationUri usando el 
        /// targetRealm. Si no se especifica ningún valor de Realm en web.config, se emitirá un desafío de autenticación al 
        /// targetApplicationUri para detectarlo.
        /// </summary>
        /// <param name="targetApplicationUri">Dirección URL del sitio de SharePoint de destino</param>
        /// <param name="identity">Identidad de Windows del usuario en cuyo nombre se va a crear el token de acceso</param>
        /// <returns>ClientContext que usa un token de acceso con una audiencia de la aplicación de destino</returns>
        public static ClientContext GetS2SClientContextWithWindowsIdentity(
            Uri targetApplicationUri,
            WindowsIdentity identity)
        {
            string realm = string.IsNullOrEmpty(Realm) ? GetRealmFromTargetUrl(targetApplicationUri) : Realm;

            JsonWebTokenClaim[] claims = identity != null ? GetClaimsWithWindowsIdentity(identity) : null;

            string accessToken = GetS2SAccessTokenWithClaims(targetApplicationUri.Authority, realm, claims);

            return GetClientContextWithAccessToken(targetApplicationUri.ToString(), accessToken);
        }

        /// <summary>
        /// Obtener dominio kerberos de autenticación de SharePoint
        /// </summary>
        /// <param name="targetApplicationUri">Dirección URL del sitio de SharePoint de destino</param>
        /// <returns>Devuelve la representación del GUID del dominio kerberos</returns>
        public static string GetRealmFromTargetUrl(Uri targetApplicationUri)
        {
            WebRequest request = WebRequest.Create(targetApplicationUri + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                if (e.Response == null)
                {
                    return null;
                }

                string bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];
                if (string.IsNullOrEmpty(bearerResponseHeader))
                {
                    return null;
                }

                const string bearer = "Bearer realm=\"";
                int bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);
                if (bearerIndex < 0)
                {
                    return null;
                }

                int realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    string targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                    {
                        return targetRealm;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Determina si se trata de un complemento de alta confianza.
        /// </summary>
        /// <returns>Verdadero si es un complemento de alta confianza.</returns>
        public static bool IsHighTrustApp()
        {
            return SigningCredentials != null;
        }

        /// <summary>
        /// Garantiza que la dirección URL especificada termina con '/', si no es NULL ni está vacía.
        /// </summary>
        /// <param name="url">Dirección URL.</param>
        /// <returns>Dirección URL que termina con '/', si no es NULL ni está vacía.</returns>
        public static string EnsureTrailingSlash(string url)
        {
            if (!string.IsNullOrEmpty(url) && url[url.Length - 1] != '/')
            {
                return url + "/";
            }

            return url;
        }

        #endregion

        #region campos privados

        //
        // Constantes de configuración
        //        

        private const string AuthorizationPage = "_layouts/15/OAuthAuthorize.aspx";
        private const string RedirectPage = "_layouts/15/AppRedirect.aspx";
        private const string AcsPrincipalName = "00000001-0000-0000-c000-000000000000";
        private const string AcsMetadataEndPointRelativeUrl = "metadata/json/1";
        private const string S2SProtocol = "OAuth2";
        private const string DelegationIssuance = "DelegationIssuance1.0";
        private const string NameIdentifierClaimType = JsonWebTokenConstants.ReservedClaims.NameIdentifier;
        private const string TrustedForImpersonationClaimType = "trustedfordelegation";
        private const string ActorTokenClaimType = JsonWebTokenConstants.ReservedClaims.ActorToken;

        //
        // Constantes de entorno
        //

        private static string GlobalEndPointPrefix = "accounts";
        private static string AcsHostUrl = "accesscontrol.windows.net";

        //
        // Configuración del complemento hospedado
        //
        private static readonly string ClientId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientId")) ? WebConfigurationManager.AppSettings.Get("HostedAppName") : WebConfigurationManager.AppSettings.Get("ClientId");
        private static readonly string IssuerId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("IssuerId")) ? ClientId : WebConfigurationManager.AppSettings.Get("IssuerId");
        private static readonly string HostedAppHostNameOverride = WebConfigurationManager.AppSettings.Get("HostedAppHostNameOverride");
        private static readonly string HostedAppHostName = WebConfigurationManager.AppSettings.Get("HostedAppHostName");
        private static readonly string ClientSecret = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientSecret")) ? WebConfigurationManager.AppSettings.Get("HostedAppSigningKey") : WebConfigurationManager.AppSettings.Get("ClientSecret");
        private static readonly string SecondaryClientSecret = WebConfigurationManager.AppSettings.Get("SecondaryClientSecret");
        private static readonly string Realm = WebConfigurationManager.AppSettings.Get("Realm");
        private static readonly string ServiceNamespace = WebConfigurationManager.AppSettings.Get("Realm");

        private static readonly string ClientSigningCertificatePath = WebConfigurationManager.AppSettings.Get("ClientSigningCertificatePath");
        private static readonly string ClientSigningCertificatePassword = WebConfigurationManager.AppSettings.Get("ClientSigningCertificatePassword");
        private static readonly X509Certificate2 ClientCertificate = (string.IsNullOrEmpty(ClientSigningCertificatePath) || string.IsNullOrEmpty(ClientSigningCertificatePassword)) ? null : new X509Certificate2(ClientSigningCertificatePath, ClientSigningCertificatePassword);
        private static readonly X509SigningCredentials SigningCredentials = (ClientCertificate == null) ? null : new X509SigningCredentials(ClientCertificate, SecurityAlgorithms.RsaSha256Signature, SecurityAlgorithms.Sha256Digest);

        #endregion

        #region métodos privados

        private static ClientContext CreateAcsClientContextForUrl(SPRemoteEventProperties properties, Uri sharepointUrl)
        {
            string contextTokenString = properties.ContextToken;

            if (String.IsNullOrEmpty(contextTokenString))
            {
                return null;
            }

            SharePointContextToken contextToken = ReadAndValidateContextToken(contextTokenString, OperationContext.Current.IncomingMessageHeaders.To.Host);
            string accessToken = GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken;

            return GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
        }

        private static string GetAcsMetadataEndpointUrl()
        {
            return Path.Combine(GetAcsGlobalEndpointUrl(), AcsMetadataEndPointRelativeUrl);
        }

        private static string GetFormattedPrincipal(string principalName, string hostName, string realm)
        {
            if (!String.IsNullOrEmpty(hostName))
            {
                return String.Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, realm);
            }

            return String.Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, realm);
        }

        private static string GetAcsPrincipalName(string realm)
        {
            return GetFormattedPrincipal(AcsPrincipalName, new Uri(GetAcsGlobalEndpointUrl()).Host, realm);
        }

        private static string GetAcsGlobalEndpointUrl()
        {
            return String.Format(CultureInfo.InvariantCulture, "https://{0}.{1}/", GlobalEndPointPrefix, AcsHostUrl);
        }

        private static JsonWebSecurityTokenHandler CreateJsonWebSecurityTokenHandler()
        {
            JsonWebSecurityTokenHandler handler = new JsonWebSecurityTokenHandler();
            handler.Configuration = new SecurityTokenHandlerConfiguration();
            handler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Never);
            handler.Configuration.CertificateValidator = X509CertificateValidator.None;

            List<byte[]> securityKeys = new List<byte[]>();
            securityKeys.Add(Convert.FromBase64String(ClientSecret));
            if (!string.IsNullOrEmpty(SecondaryClientSecret))
            {
                securityKeys.Add(Convert.FromBase64String(SecondaryClientSecret));
            }

            List<SecurityToken> securityTokens = new List<SecurityToken>();
            securityTokens.Add(new MultipleSymmetricKeySecurityToken(securityKeys));

            handler.Configuration.IssuerTokenResolver =
                SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
                new ReadOnlyCollection<SecurityToken>(securityTokens),
                false);
            SymmetricKeyIssuerNameRegistry issuerNameRegistry = new SymmetricKeyIssuerNameRegistry();
            foreach (byte[] securitykey in securityKeys)
            {
                issuerNameRegistry.AddTrustedIssuer(securitykey, GetAcsPrincipalName(ServiceNamespace));
            }
            handler.Configuration.IssuerNameRegistry = issuerNameRegistry;
            return handler;
        }

        private static string GetS2SAccessTokenWithClaims(
            string targetApplicationHostName,
            string targetRealm,
            IEnumerable<JsonWebTokenClaim> claims)
        {
            return IssueToken(
                ClientId,
                IssuerId,
                targetRealm,
                SharePointPrincipal,
                targetRealm,
                targetApplicationHostName,
                true,
                claims,
                claims == null);
        }

        private static JsonWebTokenClaim[] GetClaimsWithWindowsIdentity(WindowsIdentity identity)
        {
            JsonWebTokenClaim[] claims = new JsonWebTokenClaim[]
            {
                new JsonWebTokenClaim(NameIdentifierClaimType, identity.User.Value.ToLower()),
                new JsonWebTokenClaim("nii", "urn:office:idp:activedirectory")
            };
            return claims;
        }

        private static string IssueToken(
            string sourceApplication,
            string issuerApplication,
            string sourceRealm,
            string targetApplication,
            string targetRealm,
            string targetApplicationHostName,
            bool trustedForDelegation,
            IEnumerable<JsonWebTokenClaim> claims,
            bool appOnly = false)
        {
            if (null == SigningCredentials)
            {
                throw new InvalidOperationException("SigningCredentials was not initialized");
            }

            #region Token de actor

            string issuer = string.IsNullOrEmpty(sourceRealm) ? issuerApplication : string.Format("{0}@{1}", issuerApplication, sourceRealm);
            string nameid = string.IsNullOrEmpty(sourceRealm) ? sourceApplication : string.Format("{0}@{1}", sourceApplication, sourceRealm);
            string audience = string.Format("{0}/{1}@{2}", targetApplication, targetApplicationHostName, targetRealm);

            List<JsonWebTokenClaim> actorClaims = new List<JsonWebTokenClaim>();
            actorClaims.Add(new JsonWebTokenClaim(JsonWebTokenConstants.ReservedClaims.NameIdentifier, nameid));
            if (trustedForDelegation && !appOnly)
            {
                actorClaims.Add(new JsonWebTokenClaim(TrustedForImpersonationClaimType, "true"));
            }

            // Crear token
            JsonWebSecurityToken actorToken = new JsonWebSecurityToken(
                issuer: issuer,
                audience: audience,
                validFrom: DateTime.UtcNow,
                validTo: DateTime.UtcNow.Add(HighTrustAccessTokenLifetime),
                signingCredentials: SigningCredentials,
                claims: actorClaims);

            string actorTokenString = new JsonWebSecurityTokenHandler().WriteTokenAsString(actorToken);

            if (appOnly)
            {
                // El token solo de aplicación es el mismo que el token de actor para el caso delegado
                return actorTokenString;
            }

            #endregion Actor token

            #region Token externo

            List<JsonWebTokenClaim> outerClaims = null == claims ? new List<JsonWebTokenClaim>() : new List<JsonWebTokenClaim>(claims);
            outerClaims.Add(new JsonWebTokenClaim(ActorTokenClaimType, actorTokenString));

            JsonWebSecurityToken jsonToken = new JsonWebSecurityToken(
                nameid, // el emisor de token externo debe coincidir con el identificador de nombre del token de actor
                audience,
                DateTime.UtcNow,
                DateTime.UtcNow.Add(HighTrustAccessTokenLifetime),
                outerClaims);

            string accessToken = new JsonWebSecurityTokenHandler().WriteTokenAsString(jsonToken);

            #endregion Outer token

            return accessToken;
        }

        #endregion

        #region AcsMetadataParser

        // Esta clase se usa para obtener el documento de metadatos del extremo global de STS. Contiene
        // métodos para analizar el documento de metadatos y obtener extremos y un certificado STS.
        public static class AcsMetadataParser
        {
            public static X509Certificate2 GetAcsSigningCert(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                if (null != document.keys && document.keys.Count > 0)
                {
                    JsonKey signingKey = document.keys[0];

                    if (null != signingKey && null != signingKey.keyValue)
                    {
                        return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
                    }
                }

                throw new Exception("Metadata document does not contain ACS signing certificate.");
            }

            public static string GetDelegationServiceUrl(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                JsonEndpoint delegationEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == DelegationIssuance);

                if (null != delegationEndpoint)
                {
                    return delegationEndpoint.location;
                }
                throw new Exception("Metadata document does not contain Delegation Service endpoint Url");
            }

            private static JsonMetadataDocument GetMetadataDocument(string realm)
            {
                string acsMetadataEndpointUrlWithRealm = String.Format(CultureInfo.InvariantCulture, "{0}?realm={1}",
                                                                       GetAcsMetadataEndpointUrl(),
                                                                       realm);
                byte[] acsMetadata;
                using (WebClient webClient = new WebClient())
                {

                    acsMetadata = webClient.DownloadData(acsMetadataEndpointUrlWithRealm);
                }
                string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

                JavaScriptSerializer serializer = new JavaScriptSerializer();
                JsonMetadataDocument document = serializer.Deserialize<JsonMetadataDocument>(jsonResponseString);

                if (null == document)
                {
                    throw new Exception("No metadata document found at the global endpoint " + acsMetadataEndpointUrlWithRealm);
                }

                return document;
            }

            public static string GetStsUrl(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                JsonEndpoint s2sEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == S2SProtocol);

                if (null != s2sEndpoint)
                {
                    return s2sEndpoint.location;
                }

                throw new Exception("Metadata document does not contain STS endpoint url");
            }

            private class JsonMetadataDocument
            {
                public string serviceName { get; set; }
                public List<JsonEndpoint> endpoints { get; set; }
                public List<JsonKey> keys { get; set; }
            }

            private class JsonEndpoint
            {
                public string location { get; set; }
                public string protocol { get; set; }
                public string usage { get; set; }
            }

            private class JsonKeyValue
            {
                public string type { get; set; }
                public string value { get; set; }
            }

            private class JsonKey
            {
                public string usage { get; set; }
                public JsonKeyValue keyValue { get; set; }
            }
        }

        #endregion
    }

    /// <summary>
    /// JsonWebSecurityToken generado por SharePoint para autenticarse en una aplicación de terceros y permitir devoluciones de llamada mediante un token de actualización
    /// </summary>
    public class SharePointContextToken : JsonWebSecurityToken
    {
        public static SharePointContextToken Create(JsonWebSecurityToken contextToken)
        {
            return new SharePointContextToken(contextToken.Issuer, contextToken.Audience, contextToken.ValidFrom, contextToken.ValidTo, contextToken.Claims);
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<JsonWebTokenClaim> claims)
            : base(issuer, audience, validFrom, validTo, claims)
        {
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<JsonWebTokenClaim> claims, SecurityToken issuerToken, JsonWebSecurityToken actorToken)
            : base(issuer, audience, validFrom, validTo, claims, issuerToken, actorToken)
        {
        }

        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<JsonWebTokenClaim> claims, SigningCredentials signingCredentials)
            : base(issuer, audience, validFrom, validTo, claims, signingCredentials)
        {
        }

        public string NameId
        {
            get
            {
                return GetClaimValue(this, "nameid");
            }
        }

        /// <summary>
        /// La parte de nombre de entidad de seguridad de la notificación "appctxsender" del token de contexto
        /// </summary>
        public string TargetPrincipalName
        {
            get
            {
                string appctxsender = GetClaimValue(this, "appctxsender");

                if (appctxsender == null)
                {
                    return null;
                }

                return appctxsender.Split('@')[0];
            }
        }

        /// <summary>
        /// Notificación "refreshtoken" del token de contexto
        /// </summary>
        public string RefreshToken
        {
            get
            {
                return GetClaimValue(this, "refreshtoken");
            }
        }

        /// <summary>
        /// Notificación "CacheKey" del token de contexto
        /// </summary>
        public string CacheKey
        {
            get
            {
                string appctx = GetClaimValue(this, "appctx");
                if (appctx == null)
                {
                    return null;
                }

                ClientContext ctx = new ClientContext("http://tempuri.org");
                Dictionary<string, object> dict = (Dictionary<string, object>)ctx.ParseObjectFromJsonString(appctx);
                string cacheKey = (string)dict["CacheKey"];

                return cacheKey;
            }
        }

        /// <summary>
        /// Notificación "SecurityTokenServiceUri" del token de contexto
        /// </summary>
        public string SecurityTokenServiceUri
        {
            get
            {
                string appctx = GetClaimValue(this, "appctx");
                if (appctx == null)
                {
                    return null;
                }

                ClientContext ctx = new ClientContext("http://tempuri.org");
                Dictionary<string, object> dict = (Dictionary<string, object>)ctx.ParseObjectFromJsonString(appctx);
                string securityTokenServiceUri = (string)dict["SecurityTokenServiceUri"];

                return securityTokenServiceUri;
            }
        }

        /// <summary>
        /// Parte de dominio kerberos de la notificación "audience" del token de contexto
        /// </summary>
        public string Realm
        {
            get
            {
                string aud = Audience;
                if (aud == null)
                {
                    return null;
                }

                string tokenRealm = aud.Substring(aud.IndexOf('@') + 1);

                return tokenRealm;
            }
        }

        private static string GetClaimValue(JsonWebSecurityToken token, string claimType)
        {
            if (token == null)
            {
                throw new ArgumentNullException("token");
            }

            foreach (JsonWebTokenClaim claim in token.Claims)
            {
                if (StringComparer.Ordinal.Equals(claim.ClaimType, claimType))
                {
                    return claim.Value;
                }
            }

            return null;
        }

    }

    /// <summary>
    /// Representa un token de seguridad que contiene varias claves de seguridad que se generan mediante algoritmos simétricos.
    /// </summary>
    public class MultipleSymmetricKeySecurityToken : SecurityToken
    {
        /// <summary>
        /// Inicializa una nueva instancia de la clase MultipleSymmetricKeySecurityToken.
        /// </summary>
        /// <param name="keys">Enumeración de matrices de bytes que contienen las claves simétricas.</param>
        public MultipleSymmetricKeySecurityToken(IEnumerable<byte[]> keys)
            : this(UniqueId.CreateUniqueId(), keys)
        {
        }

        /// <summary>
        /// Inicializa una nueva instancia de la clase MultipleSymmetricKeySecurityToken.
        /// </summary>
        /// <param name="tokenId">Identificador único del token de seguridad.</param>
        /// <param name="keys">Enumeración de matrices de bytes que contienen las claves simétricas.</param>
        public MultipleSymmetricKeySecurityToken(string tokenId, IEnumerable<byte[]> keys)
        {
            if (keys == null)
            {
                throw new ArgumentNullException("keys");
            }

            if (String.IsNullOrEmpty(tokenId))
            {
                throw new ArgumentException("Value cannot be a null or empty string.", "tokenId");
            }

            foreach (byte[] key in keys)
            {
                if (key.Length <= 0)
                {
                    throw new ArgumentException("The key length must be greater then zero.", "keys");
                }
            }

            id = tokenId;
            effectiveTime = DateTime.UtcNow;
            securityKeys = CreateSymmetricSecurityKeys(keys);
        }

        /// <summary>
        /// Obtiene el identificador único del token de seguridad.
        /// </summary>
        public override string Id
        {
            get
            {
                return id;
            }
        }

        /// <summary>
        /// Obtiene las claves criptográficas asociadas al token de seguridad.
        /// </summary>
        public override ReadOnlyCollection<SecurityKey> SecurityKeys
        {
            get
            {
                return securityKeys.AsReadOnly();
            }
        }

        /// <summary>
        /// Obtiene el primer instante en el tiempo en que este token de seguridad es válido.
        /// </summary>
        public override DateTime ValidFrom
        {
            get
            {
                return effectiveTime;
            }
        }

        /// <summary>
        /// Obtiene el último instante en el tiempo en que este token de seguridad es válido.
        /// </summary>
        public override DateTime ValidTo
        {
            get
            {
                // No expira nunca
                return DateTime.MaxValue;
            }
        }

        /// <summary>
        /// Devuelve un valor que indica si el identificador de clave para esta instancia se puede resolver en el identificador de clave especificado.
        /// </summary>
        /// <param name="keyIdentifierClause">SecurityKeyIdentifierClause que se va a comparar con esta instancia.</param>
        /// <returns>Es true si keyIdentifierClause es una SecurityKeyIdentifierClause y tiene el mismo identificador único que la propiedad Id; de lo contrario, es false.</returns>
        public override bool MatchesKeyIdentifierClause(SecurityKeyIdentifierClause keyIdentifierClause)
        {
            if (keyIdentifierClause == null)
            {
                throw new ArgumentNullException("keyIdentifierClause");
            }

            // Como se trata de un token simétrico y no tenemos identificadores para distinguir los tokens, solo comprobamos la
            // presencia de un SymmetricIssuerKeyIdentifier. La asignación real al emisor se realiza más tarde
            // cuando la clave se hace coincidir con el emisor.
            if (keyIdentifierClause is SymmetricIssuerKeyIdentifierClause)
            {
                return true;
            }
            return base.MatchesKeyIdentifierClause(keyIdentifierClause);
        }

        #region miembros privados

        private List<SecurityKey> CreateSymmetricSecurityKeys(IEnumerable<byte[]> keys)
        {
            List<SecurityKey> symmetricKeys = new List<SecurityKey>();
            foreach (byte[] key in keys)
            {
                symmetricKeys.Add(new InMemorySymmetricSecurityKey(key));
            }
            return symmetricKeys;
        }

        private string id;
        private DateTime effectiveTime;
        private List<SecurityKey> securityKeys;

        #endregion
    }
}
