Buscador de stock v3.0.0 - Portainer

Archivos principales:
- stock-buscador.html + app.js: interfaz (centro de control, tema oscuro/claro).
- server.py + stock_processor.py: API, historial y preprocesado.
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
