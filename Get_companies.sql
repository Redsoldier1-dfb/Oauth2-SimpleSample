DECLARE @Object INT;
DECLARE @contentType NVARCHAR(255);
DECLARE @Url NVARCHAR(2048) = 'https://api.businesscentral.dynamics.com/v2.0/{Idtenant}/Sandbox/api/v2.0/companies';
DECLARE @Token NVARCHAR(2048) = 'Bearer {TOKEN}';  -- Si es necesario
DECLARE @ResponseText AS TABLE(ResponseText NVARCHAR(MAX));

-- Crear un objeto HTTP
EXEC sp_OACreate 'MSXML2.ServerXMLHTTP', @Object OUT;

-- Abrir una conexi�n
EXEC sp_OAMethod @Object, 'open', NULL, 'GET', @url, 'false';

-- Establecer el encabezado de autorizaci�n
EXEC sp_OAMethod @Object, 'setRequestHeader', NULL, 'Accept', 'application/json;odata.metadata=none';

-- Establecer el encabezado de quitar meta data
EXEC sp_OAMethod @Object, 'setRequestHeader', NULL, 'Authorization', @token;

-- Enviar la solicitud
EXEC sp_OAMethod @Object, 'send', NULL;

-- Obtener el tipo de contenido
EXEC sp_OAGetProperty @Object, 'getResponseHeader', @contentType OUT, 'Content-Type';

-- Insert response into @ResponseText Table
INSERT INTO @ResponseText (ResponseText) EXEC sp_OAGetProperty @Object, 'responseText'
Select * from @ResponseText

-- Limpiar
EXEC sp_OADestroy @Object;
