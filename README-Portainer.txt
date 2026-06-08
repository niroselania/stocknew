Buscador de stock - Portainer

Archivos principales:
- stock-buscador.html + app.js: interfaz.
- server.py + stock_processor.py: API, historial y preprocesado.
- xlsx-worker.js: lectura de Excel en segundo plano en el navegador.
- Dockerfile: imagen Python con openpyxl.

Stack recomendado en Portainer:
1. Subí esta carpeta a Git o al servidor.
2. En Portainer: Stacks > Add stack.
3. Usá este compose:

services:
  stock-buscador:
    build: .
    container_name: stock-buscador
    restart: unless-stopped
    ports:
      - "8088:8000"
    volumes:
      - stock-data:/data
    environment:
      HISTORY_MAX: "30"

volumes:
  stock-data:

4. Deploy the stack.
5. Abrilo en http://IP-DE-TU-SERVIDOR:8088

Si ya tenías la versión anterior con volumen en /opt/stock-buscador/data,
podés usar docker-compose-portainer-con-planilla.yml.

Novedades:
- Historial de planillas y comparación de stock entre cargas.
- Carga más rápida usando datos preprocesados en el servidor.
- Filtros en stock completo, exportar resultado, escáner, PWA y modo oscuro.

Notas:
- Si el puerto 8088 está ocupado, cambialo por otro, por ejemplo "8090:8000".
- HISTORY_MAX controla cuántas planillas viejas se conservan (default 30).
- La primera vez que arranque con una planilla vieja, el servidor la reprocesa solo.
