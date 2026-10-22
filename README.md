# rifa-boletas-admin
Prueba de Registro de Boletas en HTML y Google Docs

# Publicar el administrador en GitHub Pages

## 1. Preparar el proyecto

1. Cree un repositorio nuevo en GitHub, por ejemplo `rifa-boletas-admin`.
2. Copie en el repositorio:
   - `Administrar_Boletas_GoogleSheets.html`
   - opcionalmente renombre el archivo a `index.html`
3. Edite la constante:

```javascript
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/ID_DEPLOYMENT/exec';
```

Reemplace `ID_DEPLOYMENT` por la URL de su Web App de Google Apps Script.


## 2. Configurar Google Sheets y Apps Script

1. Abra el Google Sheet vinculado a la rifa.
2. Vaya a **Extensiones > Apps Script**.
3. Pegue el contenido de `GoogleAppsScript_Code.gs`.
4. Guarde el proyecto.
5. Haga clic en **Implementar > Nueva implementación**.
6. Tipo: **Aplicación web**.
7. Ejecutar como: **Yo**.
8. Quién tiene acceso: **Cualquier persona**.
9. Copie la URL `/exec`.
10. Pegue esa URL en `APPS_SCRIPT_URL` del HTML.

## 3. Publicar en GitHub Pages

1. Suba los archivos al repositorio.
2. Vaya a **Settings > Pages**.
3. En **Source**, seleccione:
   - **Deploy from a branch**
   - Branch: `main`
   - Folder: `/ (root)`
4. Guarde.
5. Espere la URL pública, normalmente:

```text
https://usuario.github.io/rifa-boletas-admin/
```

## 4. Probar

1. Abra la URL pública.
2. Seleccione un asociado.
3. Busque una boleta por número.
4. Actualice los datos.
5. Haga clic en **Guardar cambios**.
6. Verifique en Google Sheets que la boleta y la hoja del socio se actualizaron.

## 5. Seguridad

- No ponga credenciales de Google dentro del HTML.
- La autenticación y los permisos se manejan desde Apps Script.
- Si la hoja contiene datos sensibles, restrinja el acceso de la Web App según su necesidad.