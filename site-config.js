/* Configuración de BAT 360 Merch.
   - usuarios: correos con acceso y qué vistas ven ('*' = todas). Para dar acceso a otra persona: agrega su correo.
   - vistas: entradas del menú lateral. Para sumar una vista nueva: sube su .html a esta carpeta y agrega una línea aquí. */
window.SITE = {
  id: 'merch',
  nombre: 'Merch',
  // Mismo Client ID de Google que BAT 360. OJO: este dominio nuevo debe agregarse en Google Cloud → "Orígenes JavaScript autorizados".
  googleClientId: '382427877987-j68n1tmm6ers3biekqamem2jlrdguubj.apps.googleusercontent.com',
  usuarios: {
    'ccamarena@ttaudit.com': { vistas: ['*'] },
    'mcarhuallanqui@ttaudit.com': { vistas: ['*'] }
  },
  vistas: [
    { id: 'consolidado', seccion: 'Merch', icono: '📋', titulo: 'Consolidado',
      sub: 'Órdenes de trabajo del consolidado (hoja NUEVO 2222): programadas, reportadas, efectivas y no efectivas por proyecto, equipo y distrito.',
      src: 'consolidado.html' },
    { id: 'proyectos', seccion: 'Merch', icono: '📸', titulo: 'Proyectos',
      sub: 'Avance de cada campaña con fotos de las tiendas. Excluye HW y Proyecto Interno (tienen vistas propias).',
      src: 'proyectos.html' },
    { id: 'inicio', seccion: 'Merch', icono: '🏠', titulo: 'Inicio',
      sub: 'Espacio de Merch de BAT 360. Aquí se irán sumando las vistas.',
      src: 'inicio.html' }
  ]
};
