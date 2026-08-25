# EAI Bolivia Economic Monitor — Cloud Run

Rama de despliegue del **Bolivia Economic Monitor** de Economic Analytics Institute (EAI).

## Objetivo

Mantener el monitor disponible de forma continua en Google Cloud Run, evitando la hibernación de Streamlit Community Cloud.

## Arquitectura

EAI Website (Cloudflare) → Bolivia Economic Monitor (Cloud Run) → Google Drive/Excel

## Fuente de datos identificada

Se configuró por defecto el archivo de Drive **Info.xlsx** con ID:

`12UDMaTEeqvNDaFH8sUDJy9E6YYLehI_J`

Antes de desplegar, comparte ese archivo con la cuenta de servicio que crea el script, con permiso **Lector/Viewer**.

## Seguridad

La versión Cloud Run usa **Application Default Credentials (ADC)**. No requiere guardar una llave JSON de cuenta de servicio en GitHub, Streamlit ni variables de entorno.

## Despliegue en Windows

1. Instala Google Cloud CLI e inicia sesión con `gcloud auth login`.
2. Crea o selecciona un proyecto de Google Cloud con facturación habilitada.
3. Abre PowerShell en la carpeta del repositorio, en la rama `eai-cloud-run`.
4. Ejecuta:

```powershell
./deploy_cloud_run.ps1 -ProjectId "TU_PROJECT_ID"
```

El script:

- habilita Cloud Run, Cloud Build, Artifact Registry, IAM y Drive API;
- crea la cuenta `eai-monitor@TU_PROJECT_ID.iam.gserviceaccount.com`;
- te pide compartir `Info.xlsx` con esa cuenta;
- despliega el servicio público;
- configura `GOOGLE_DRIVE_FILE_ID` y `EAI_DRIVE_AUTH=adc`;
- establece 1 instancia mínima y 1 máxima;
- muestra la URL pública final.

## Configuración inicial

- Región: `southamerica-west1` (Santiago)
- Servicio: `eai-bolivia-economic-monitor`
- CPU: 1
- Memoria: 1 GiB
- Instancias mínimas: 1
- Instancias máximas: 1
- Timeout: 3600 s
- Acceso: público

La instancia mínima evita escalar a cero. El máximo de una instancia mantiene las sesiones de Streamlit en un solo proceso mientras el tráfico sea moderado. Esta configuración se puede ampliar posteriormente.

## Cómo se genera la versión EAI

Para no modificar el dashboard institucional de la rama `main`, el `Dockerfile` toma el `app.py` original y aplica durante la construcción un parche ubicado en `deploy/eai_patch_*`. El contenedor resultante muestra únicamente la identidad EAI y habilita autenticación ADC para Google Drive.

## Próximo paso

Cuando Cloud Run entregue la URL `https://...run.app`, enlazarla desde la sección **Monitor** del sitio principal de Economic Analytics Institute. Luego puede configurarse un subdominio propio a través de Cloudflare.
