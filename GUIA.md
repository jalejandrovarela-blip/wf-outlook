# Guía de instalación: botón "Llenar cuerpo WF" (nuevo Outlook)

## Paso 1. Publicar los archivos (una sola vez, ~10 min)
Los complementos de Outlook necesitan estar en una dirección HTTPS. La opción más simple y gratuita es GitHub Pages:

1. Crea una cuenta en github.com (si no tienes).
2. Crea un repositorio **público** llamado `wf-outlook`.
3. Sube a él: `commands.html`, `commands.js`, `manifest.xml` y la carpeta `assets` (con los iconos). Usa "Add file > Upload files".
4. En el repositorio: Settings > Pages > Source: "Deploy from a branch" > rama `main`, carpeta `/ (root)` > Save.
5. Espera 1-2 minutos. Tu URL será `https://jalejandrovarela-blip.github.io/wf-outlook/`. Compruébala abriendo `.../commands.html` en el navegador (debe cargar en blanco, sin error 404).

> Los archivos no contienen datos sensibles, solo la tabla OSCAC/CAAT. Si prefieres no usar un repositorio público, puedes alojarlos en SharePoint o en un sitio web de Invermesa, siempre que sea HTTPS.

## Paso 2. Poner tu URL en el manifiesto
Abre `manifest.xml` con el Bloc de notas y reemplaza **todas** las apariciones de `jalejandrovarela-blip` por tu usuario de GitHub (Ctrl+H, "Reemplazar todo"). Guarda el archivo en tu computadora (no hace falta subirlo de nuevo).

## Paso 3. Instalarlo en Outlook
1. En el nuevo Outlook, abre un **correo nuevo**.
2. En la cinta, pulsa **⋯ (más opciones) > Aplicaciones** (o "Obtener complementos").
3. Elige **Mis complementos > Complementos personalizados > Agregar un complemento personalizado > Agregar desde archivo**, y selecciona `manifest.xml`.
4. Acepta el aviso. Se instalará en tu cuenta y quedará disponible siempre.

Si esa opción no aparece o dice que está bloqueada, tu administrador de Microsoft 365 la restringió; pídele que permita complementos personalizados o que lo implemente desde el Centro de administración (Configuración > Aplicaciones integradas > Cargar aplicaciones personalizadas).

## Paso 4. Uso diario
1. Correo nuevo, asunto `6943/3443/1045 Exportación Invermesa Nogales`.
2. En PARA, la lista de distribución (Wonder Fields Nogales, Wonder Fields McAllen o Agencia Aduanal McAllen Lipman).
3. Pulsa el botón **Llenar cuerpo WF** (grupo "Wonder Fields", en la pestaña Mensaje o dentro de ⋯ Aplicaciones). Un aviso confirmará los datos puestos.
4. Adjunta tus archivos y envía. La firma no se toca.

Si el cuerpo ya tiene "CRUCE MAÑANA", no lo duplica. Si el asunto no tiene el formato correcto o la lista no se reconoce, te lo avisa en lugar de llenar algo equivocado.

## Cómo modificar datos después
Edita `commands.js` (arriba, tabla `AGENCIAS`) y vuelve a subirlo a GitHub. No necesitas reinstalar. A veces Outlook guarda en caché la versión anterior; reinicia Outlook si no ve el cambio.

- Nueva agencia: agrega una línea con el nombre de la lista, OSCAC y CAAT.
- Fecha que salte fines de semana: cambia `SALTAR_FIN_DE_SEMANA` a `true`.
- CAJA y MEDIDA: constantes al inicio del archivo.

## Datos configurados
| Lista en PARA | OSCAC | CAAT |
|---|---|---|
| Wonder Fields Nogales (también "Agencia Aduanal Nogales Wonder") | PYSR | 32VT |
| Wonder Fields McAllen | PJDM | 3J50 |
| Agencia Aduanal McAllen Lipman | PJDM | 3J50 |
