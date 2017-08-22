(function (window, undefined) {

    "use strict";

    var $ = window.jQuery;
    var document = window.document;

    // Nombre del parámetro SPHostUrl
    var SPHostUrlKey = "SPHostUrl";

    // Obtiene el parámetro SPHostUrl de la dirección URL actual y lo anexa como cadena de consulta a los vínculos que señalan al dominio actual en la página.
    $(document).ready(function () {
        ensureSPHasRedirectedToSharePointRemoved();

        var spHostUrl = getSPHostUrlFromQueryString(window.location.search);
        var currentAuthority = getAuthorityFromUrl(window.location.href).toUpperCase();

        if (spHostUrl && currentAuthority) {
            appendSPHostUrlToLinks(spHostUrl, currentAuthority);
        }
    });

    // Anexa SPHostUrl como cadena de consulta a todos los vínculos que señalan al dominio actual.
    function appendSPHostUrlToLinks(spHostUrl, currentAuthority) {
        $("a")
            .filter(function () {
                var authority = getAuthorityFromUrl(this.href);
                if (!authority && /^#|:/.test(this.href)) {
                    // Filtra los delimitadores y las direcciones URL con otros protocolos no admitidos.
                    return false;
                }
                return authority.toUpperCase() == currentAuthority;
            })
            .each(function () {
                if (!getSPHostUrlFromQueryString(this.search)) {
                    if (this.search.length > 0) {
                        this.search += "&" + SPHostUrlKey + "=" + spHostUrl;
                    }
                    else {
                        this.search = "?" + SPHostUrlKey + "=" + spHostUrl;
                    }
                }
            });
    }

    // Obtiene el parámetro SPHostUrl de la cadena de consulta dada.
    function getSPHostUrlFromQueryString(queryString) {
        if (queryString) {
            if (queryString[0] === "?") {
                queryString = queryString.substring(1);
            }

            var keyValuePairArray = queryString.split("&");

            for (var i = 0; i < keyValuePairArray.length; i++) {
                var currentKeyValuePair = keyValuePairArray[i].split("=");

                if (currentKeyValuePair.length > 1 && currentKeyValuePair[0] == SPHostUrlKey) {
                    return currentKeyValuePair[1];
                }
            }
        }

        return null;
    }

    // Obtiene la autoridad de la dirección URL especificada cuando se trata de una dirección URL absoluta con protocolo HTTP/HTTPS o una dirección URL relativa de protocolo.
    function getAuthorityFromUrl(url) {
        if (url) {
            var match = /^(?:https:\/\/|http:\/\/|\/\/)([^\/\?#]+)(?:\/|#|$|\?)/i.exec(url);
            if (match) {
                return match[1];
            }
        }
        return null;
    }

    // Si SPHasRedirectedToSharePoint existe en la cadena de consulta, quítelo.
    // Así, cuando el usuario agregue marcadores a la dirección URL, SPHasRedirectedToSharePoint no se incluirá.
    // Tenga en cuenta que la modificación de window.location.search generará una solicitud adicional para el servidor.
    function ensureSPHasRedirectedToSharePointRemoved() {
        var SPHasRedirectedToSharePointParam = "&SPHasRedirectedToSharePoint=1";

        var queryString = window.location.search;

        if (queryString.indexOf(SPHasRedirectedToSharePointParam) >= 0) {
            window.location.search = queryString.replace(SPHasRedirectedToSharePointParam, "");
        }
    }

})(window);
