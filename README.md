# Almuerzo del sábado 10-10-26

Página para que cada invitado registre su pedido (Sopita, Principio, Proteína, Jugo y Observación).
Los pedidos llegan a una hoja de cálculo de Google con dos pestañas: **Resumen para la cocina** y **Pedido por invitado**.

## Conectar la hoja de Google (una sola vez, unos 3 minutos)
1. Abre https://sheets.new (crea una hoja nueva) y ponle de nombre "Almuerzo sábado 10-10-26".
2. Menú **Extensiones → Apps Script**. Borra todo lo que aparece y pega el contenido de `apps-script/Codigo.gs`. Pulsa el ícono de guardar.
3. En la barra de arriba elige la función **setup** y pulsa **Ejecutar**. Google pedirá permisos: *Revisar permisos → tu cuenta → Configuración avanzada → Ir a (no seguro) → Permitir*. Es tu propio script.
   Vuelve a la hoja: ya verás las dos pestañas con el diseño de la página.
4. En Apps Script: **Implementar → Nueva implementación** → ícono de engranaje → **Aplicación web**.
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier persona**
   - **Implementar** y copia la URL que termina en `/exec`.
5. Esa URL va en `config.js`: `window.PEDIDOS_API = "https://script.google.com/macros/s/…/exec";`

## Publicar en GitHub Pages
En el repositorio: **Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: main / (root) → Save**.
La página queda en `https://bartels-franz.github.io/almuerzo/`.
