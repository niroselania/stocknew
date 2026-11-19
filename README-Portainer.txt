Buscador de stock v3.1.1 - Portainer

Archivos principales:
- stock-buscador.html + app.js: interfaz (centro de control, tema oscuro/claro).
- server.py + stock_processor.py: API, historial y preprocesado.
- telegram_bot.py: bot de Telegram (corre dentro del mismo servidor).
- xlsx-worker.js: lectura de Excel en el navegador (respaldo si no hay servidor).
- Dockerfile: imagen Python con openpyxl.

Actualizar desde la v2.1:
1. Reemplazá los archivos del repositorio/carpeta por estos.
2. En Portainer: Stacks > tu stack > "Update the stack" con rebuild / re-pull activado.
3. No hace falta borrar el volumen: al arrancar, el servidor convierte solo la
   planilla actual al formato nuevo y el historial anterior se conserva.
4. Si algún navegador sigue mostrando la pantalla vieja, recargá con Ctrl+F5
   (solo la primera vez; después la app se actualiza sola).

Instalación nueva (Stacks > Add stack):

services:
  stock-buscador:
    build: .
    container_name: stock-buscador
    restart: unless-stopped
    ports:
      - "8085:8000"
    volumes:
      - stock-data:/data
    environment:
      HISTORY_MAX: "30"
      # Opcional: pedir un código para subir o borrar planillas
      # UPLOAD_TOKEN: "cambiar-este-codigo"

volumes:
  stock-data:

Abrilo en http://IP-DE-TU-SERVIDOR:8085

BOT DE TELEGRAM (v3.1.1)
Mandás un código y responde con código, color, talle, cantidad y precio de las
variantes que tienen stock en LOCAL. Si no hay unidades en LOCAL responde
"no hay stock en LOCAL". Si el código no existe, lo avisa.
  25551          -> todas las variantes del SKU con stock en LOCAL
  25551 NLSA     -> sólo ese color
  25551 NLSA M   -> esa variante
  /stock 25551   -> lo mismo (útil en grupos)
  /id            -> muestra el id del chat
  /ayuda         -> instrucciones

Activarlo:
1. En Telegram hablale a @BotFather, mandá /newbot, elegí nombre y usuario.
   Te da un token tipo 123456789:AAxxxx...  (no lo compartas).
2. En Portainer > Stacks > tu stack > Environment variables agregá:
     TELEGRAM_BOT_TOKEN   = el-token
   (el docker-compose.yml ya las pasa al contenedor con ${TELEGRAM_BOT_TOKEN}).
   Guardá con "Redeploy" tildado.
3. Escribile /id al bot desde tu chat (o desde un grupo donde lo agregues).
   Agregá otra variable:
     TELEGRAM_ALLOWED_CHATS = 123456789,-1001234567890
   (varios separados por coma) y volvé a hacer Redeploy.
   Sin esta lista cualquiera que encuentre el bot puede consultar stock y precios.
4. Verificación: http://IP-DE-TU-SERVIDOR:8085/api/health muestra
   "telegram": {"connected": true, "username": "..."}.
No hace falta abrir puertos: el servidor sale a Telegram (necesita internet).
No uses el mismo token en dos contenedores a la vez.

Precio: se toma ÚNICAMENTE de la columna de precio RETAIL (encabezado "RETAIL",
"PRECIO RETAIL" o que contenga la palabra RETAIL), en la pestaña STOCK o en una
pestaña PRECIOS / LISTA DE PRECIOS (con CONCAT o SKU). Otras columnas de precio
se ignoran. La columna "RE" es el stock de la base RETAIL y nunca se usa como precio.
Si una prenda no tiene valor en RETAIL, el bot muestra "sin precio".
Si el encabezado se llama distinto, definí PRICE_COLUMN con el nombre exacto.
Al subir cada planilla, el registro del contenedor indica qué columna de precio usó.

Novedades v3.0.0
Interfaz
- Diseño nuevo tipo "centro de control": tema oscuro tecnológico (y claro opcional),
  indicadores animados, barras por base, consola de estado.
- Buscador con sugerencias en vivo: por CONCAT, SKU o nombre, palabras en
  cualquier orden, flechas y Enter. Atajo: tecla / o Ctrl+K.
- Ficha de producto: total de unidades, barras por base y las otras variantes
  (colores/talles) del mismo SKU con su stock, clickeables.
- Búsquedas recientes y enlace para compartir una consulta (?q=CODIGO).
- Subir planilla arrastrándola a la pantalla, con barra de progreso real.
- Stock completo: orden por columna, filtros por estado (con stock / sin stock /
  con negativos), clic en una fila abre el producto, exportar la vista a CSV
  (compatible con Excel en español).
- Comparación: unidades que suben/bajan, productos nuevos/eliminados,
  balance por base, filtros y exportación CSV.
- Historial: descargar cualquier planilla anterior y eliminar las que sobran.
- En el celular la consulta aparece primero y la navegación va abajo.
- Funciona sin conexión con los últimos datos descargados.

Correcciones
- Planillas con columnas de fecha hacían fallar el procesado en el servidor.
- La lectura de Excel en el navegador (respaldo) no funcionaba.
- Si la fila TOTAL viene vacía, los totales se calculan sumando.
- Después de un redeploy algunos navegadores seguían con la versión vieja.
- Dos cargas en el mismo segundo podían pisarse en el historial.

Rendimiento
- Los datos que baja el navegador son unas 20 veces más livianos (sin datos
  duplicados + compresión) y no se vuelven a bajar si la planilla no cambió.
- Tabla y búsqueda optimizadas para planillas grandes.

Notas
- Si el puerto 8085 está ocupado, cambialo, por ejemplo "8090:8000".
- HISTORY_MAX controla cuántas planillas se conservan (default 30).
- UPLOAD_TOKEN (opcional): si lo definís, subir o borrar planillas pide ese
  código (se pide una vez por navegador). Consultar no pide nada.

SI EL BOT NO RESPONDE (v3.1.2)
1. Abrí http://IP:8085/api/health y mirá "telegram":
   - enabled:false  -> el contenedor no recibió TELEGRAM_BOT_TOKEN (cargalo en Portainer > Environment variables y Update/Redeploy).
   - connected:false + lastError -> el contenedor no llega a Telegram (internet/DNS) o el token es inválido.
   - products:0 -> falta subir la planilla a este servidor.
   - allowedChats:0 -> cualquiera puede consultar; si pusiste ids y responde "no autorizado", revisá TELEGRAM_ALLOWED_CHATS.
2. No uses el mismo token en dos contenedores (error 409).
3. En grupos, escribí /stock CODIGO (Telegram no entrega texto común a bots con privacidad activa).
4. IMPORTANTE: subí a GitHub TODOS los archivos (incluido telegram_bot.py, Dockerfile y docker-compose.yml nuevos).
