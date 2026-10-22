# VEMO IA · comentarios MFR

Herramienta local de FP&A: lee el FY Financials, las notas de celda, Dynamics, los decks MFR publicados y las juntas, y redacta los boxes del MFR (vs. Budget y MoM) con un resumen de la empresa.

## Uso (Windows)

- `VEMO_IA.bat`: menú en consola (resumen + comentarios, preguntar, aprender, dashboard).
- `VEMO_DASHBOARD.bat`: abre el dashboard en el navegador (`http://127.0.0.1:8765`, solo local).

Requiere Python 3.10+ (`openpyxl`, `python-pptx`; los .bat los instalan).

## Datos

Este repositorio contiene solo código. Los datos (Excel, decks, juntas, `datos/`, `salidas/`) viven en OneDrive y no se suben.

## Dynamics 365 Business Central (API)

`VEMO_IA.bat` → opción **5 › a) configurar** (o `python vemo_ia.py dynamics --configurar`). Pega el link OData del servicio, el Client ID y el secreto. El secreto se guarda en el Administrador de credenciales de Windows, nunca en OneDrive ni en este repo. Con "solo API" ya no se usa el `DB.xlsx`.

Lo que debe darte TI (una sola vez):

1. Una app registrada en Microsoft Entra ID (Client ID + Client Secret) con permiso de aplicación **Dynamics 365 Business Central › API.ReadWrite.All** y consentimiento de administrador.
2. En Business Central, página **Microsoft Entra Applications**: agregar ese Client ID, estado *Enabled*, con un permission set de solo lectura sobre los movimientos contables.

Alternativa sin secreto: app pública (“Allow public client flows”) con permiso delegado **Financials.ReadWrite.All**; eliges la opción 2 e inicias sesión con tu cuenta.

Uso: el dashboard tiene el botón **Actualizar Dynamics**; cada corrida refresca los últimos meses (incremental) y, si la API falla, usa lo último que bajó y lo avisa.
