# BAT 360 Merch

Web independiente de Merch (separada de BAT 360). Por ahora solo tiene la vista **Inicio**: es la base para ir sumando vistas.

## Archivos
- `index.html` — shell: login con Google + menú lateral + una vista por entrada (no se edita).
- `site-config.js` — **aquí se cambia todo**: usuarios con acceso y vistas del menú.
- `inicio.html` — vista de bienvenida.

## Agregar una vista nueva
1. Subir su página `.html` a esta carpeta.
2. Agregar una entrada en `vistas` de `site-config.js`.

## Publicar (una sola vez)
1. Crear el repositorio en GitHub y subir esta carpeta.
2. En Vercel: *Add New → Project* → importar el repo (sin build, es estático).
3. **Google Cloud** → Credenciales → el Client ID de BAT 360 → *Orígenes JavaScript autorizados* → agregar el dominio nuevo de Vercel.
4. Si una vista lee Excel de SharePoint con login Microsoft: **Entra ID** → App registrations → bat360 → Authentication → *Single-page application* → agregar la página como Redirect URI (`https://<dominio-nuevo>/<pagina>.html`).
