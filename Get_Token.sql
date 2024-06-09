DECLARE @Object INT;
DECLARE @ResponseText AS TABLE(ResponseText NVARCHAR(MAX));
DECLARE @Url NVARCHAR(MAX) = 'https://login.microsoftonline.com/{Idtenant}/oauth2/v2.0/token';
DECLARE @ClientId NVARCHAR(MAX) = '{Idappcli}';
DECLARE @ClientSecret NVARCHAR(MAX) = '{Secreto}';
DECLARE @Scope NVARCHAR(MAX) = 'https://api.businesscentral.dynamics.com/.default';
DECLARE @Body NVARCHAR(MAX);
DECLARE @ContentType NVARCHAR(100);

-- Crear un objeto HTTP
EXEC sp_OACreate 'MSXML2.ServerXMLHTTP', @Object OUT;

-- Abrir una conexi�n
EXEC sp_OAMethod @Object, 'open', NULL, 'POST', @Url, 'false';

-- Establecer el encabezado de Content-Type
EXEC sp_OAMethod @Object, 'setRequestHeader', NULL, 'Content-Type', 'application/x-www-form-urlencoded';

-- Preparar el cuerpo de la solicitud
SET @Body = 'grant_type=client_credentials'
    + '&client_id=' + @ClientId
    + '&client_secret=' + @ClientSecret
    + '&scope=' + @Scope;

-- Enviar la solicitud
EXEC  sp_OAMethod @Object, 'send', NULL, @Body;

-- Obtener el tipo de contenido
EXEC  sp_OAGetProperty @Object, 'getResponseHeader', @ContentType OUT, 'Content-Type';

-- Insertar la respuesta en la tabla @ResponseText
INSERT INTO @ResponseText (ResponseText)
EXEC sp_OAGetProperty @Object, 'responseText';

-- Mostrar la respuesta
SELECT * FROM @ResponseText;

-- Limpiar
EXEC sp_OADestroy @Object;


