# ChokeHound - BloodHound CE Choke Points Analyzer

<p align="center">
  <img src="logo.png" alt="ChokeHound logo" width="240">
</p>

ChokeHound analyzes BloodHound CE data to highlight choke points—privileged edges where non–Tier‑0 principals can directly impact Tier‑0 assets. The generated Excel report and optional risk log help defenders prioritize remediation.

## Features

- Identify (critical) direct non–Tier‑0 → Tier‑0 privilege relationships.
- Prioritize choke points using a weighted risk score.
- Measure how many attack paths are affected per choke point.
- Export Excel report.
- Optionally produce a detailed text log explaining each risk calculation.

## Prerequisites

- Python 3.7 or later.
- Neo4j 4.x+ accessible to the script.
- BloodHound CE data already ingested into Neo4j.
- Network access to the BloodHound Neo4j instance.

## Requirements Before Execution

The following steps are required for ChokeHound to be able to produce meaningful results before executing it:

1. **BloodHound CE installation**: Deploy BloodHound CE and confirm the Neo4j backend is initialized and reachable with valid credentials. [Instructions here](https://bloodhound.specterops.io/get-started/quickstart/community-edition-quickstart).
2. **Data collection with SharpHound or AzureHound**: Run the official collectors and import the output into BloodHound CE: [SharpHound](https://bloodhound.specterops.io/collect-data/ce-collection/sharphound) on-prem / [AzureHound](https://bloodhound.specterops.io/collect-data/ce-collection/azurehound) for Entra ID.
3. **Tier 0 configuration inside BloodHound**: Tag every Tier‑0 object with `Tag_Tier_Zero` so ChokeHound can differentiate privileged targets.
   - In BloodHound CE -> Administration -> Configuration -> Early Access Features -> Enable Tier Management Engine​
   - Add your Tier 0 oject in section "Privilege Zone Management". Tier 0 members to consider: [Tier Zero: Members and Modification](https://bloodhound.specterops.io/get-started/security-boundaries/tier-zero-members)
​

## Installation

1. **Clone the repository**  
   ```bash
   git clone https://github.com/ProtAAPP/ChokeHound.git
   cd ChokeHound
   ```
2. **Create and activate a virtual environment**
   - Windows (PowerShell)
     ```powershell
     python -m venv .venv
     .\.venv\Scripts\Activate.ps1
     ```
   - Linux / macOS
     ```bash
     python3 -m venv .venv
     source .venv/bin/activate
     ```
3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```
4. **Verify Neo4j availability**  
   Confirm your BloodHound Neo4j service is running before executing the script.

## Configuration

Adjust `config.py` to match your Neo4j endpoint (only if different from default):

```python
NEO4J_URI = "bolt://localhost:7687"
NEO4J_USER = "neo4j"
NEO4J_PASSWORD = "bloodhoundcommunityedition"
DEFAULT_OUTPUT_FILENAME = "ChokeHound_report.xlsx"
LIMIT_CHOKE_POINTS = 200
```

## Usage

ChokeHound uses CLI arguments for output customization and logging.

## Basic Usage

```bash
python ChokeHound.py
```

Generates `ChokeHound_report.xlsx` using default settings.

## Specify Output File

```bash
python ChokeHound.py --output my_report.xlsx
```

Creates a report with a custom filename (the `.xlsx` extension is appended automatically).

## Generate Risk Calculation Log

```bash
python ChokeHound.py --output my_report.xlsx --log-risk
```

Produces both the Excel workbook and `my_report_risk_calculation_log.txt`, which documents each choke point’s scoring breakdown.

## Risk Calculation for Choke Points

ChokeHound calculates a weighted risk score per choke point:

```
Risk = (SourceObjectWeight × SourceObjectCategory) +
       (RelationshipTypeWeight × RelationshipTypeCategory) +
       (TargetObjectWeight × TargetObjectCategory) +
       (AffectedAttackPathsWeight × PathsMultiplier × 10)
```

The resulting risk is normalized to a 1–100 scale, ensuring consistent comparison across findings.

_Note: The scoring model is still evolving. Review and adapt the weights and category values in `risk_config.py` for each Active Directory environment before relying on the scores._


## Customization

Tailor the scoring model and query limits via the config files.

### Adjust Risk Weights

Update the `RISK_WEIGHTS` dictionary in `risk_config.py` to emphasize or de-emphasize specific components.

### Modify Risk Categories

Edit `SOURCE_OBJECT_CATEGORIES`, `RELATIONSHIP_TYPE_CATEGORIES`, and `TARGET_OBJECT_CATEGORIES` in `risk_config.py` to reflect your threat model.

### Change Result Limit

Set `LIMIT_CHOKE_POINTS` in `config.py` to control how many prioritized choke points are included in the Excel output.

## TODO

- Add Azure / Entra ID support.
- Analyze and estimate risk for other choke points that do not directly hit Tier 0.
- Keep improving and tuning the risk estimation model.
- Configure additional tiers (not only Tier 0).
- Add recommendations to mitigate detected Choke Points.

## References

[BloodHound Community Edition Quickstart](https://bloodhound.specterops.io/get-started/quickstart/community-edition-quickstart)

[Tier Zero: Members and Modification](https://bloodhound.specterops.io/get-started/security-boundaries/tier-zero-members)

[Microsoft Enterprise Access Model](https://learn.microsoft.com/en-us/security/privileged-access-workstations/privileged-access-access-model)


## Author

[@ciyinet](https://x.com/ciyinet)

[@gobispo](https://x.com/gobispo)


# ChokeHound - Analizador de Choke Points con BloodHound CE

ChokeHound analiza datos de BloodHound CE para resaltar Choke Points: relaciones de privilegio donde objetos no Tier‑0 impactan directamente activos Tier‑0. El informe Excel y el log opcional permiten priorizar la mitigación.

## Características

- Identifica relaciones (críticas) directas entre objetos no Tier‑0 y Tier‑0.
- Prioriza cada Choke Point mediante una puntuación de riesgo ponderada.
- Mide cuántas rutas de ataque se ven afectadas.
- Generación de informe en Excel.
- Opcionalmente genera un log textual con el detalle de cada cálculo del riesgo.

## Requisitos Previos

- Python 3.7 o superior.
- Neo4j 4.x+ accesible para el script.
- Datos de BloodHound CE ya importados en Neo4j.
- Acceso de red a la instancia de BloodHound.

## Requisitos Antes de la Ejecución

Para que ChokeHound pueda producir resultados significativos, se requieren los siguientes pasos antes de su ejecución:

1. **Instalación de BloodHound CE**: Despliega BloodHound CE y confirma que Neo4j está inicializado y accesible con credenciales válidas. [Instrucciones aquí](https://bloodhound.specterops.io/get-started/quickstart/community-edition-quickstart). 
2. **Recolección de datos con SharpHound o AzureHound**: Ejecuta los colectores de datos oficiales y sube los resultados a BloodHound CE. [SharpHound](https://bloodhound.specterops.io/collect-data/ce-collection/sharphound) on-prem / [AzureHound](https://bloodhound.specterops.io/collect-data/ce-collection/azurehound) para Entra ID.
3. **Configuración de Tier 0 en BloodHound**: Etiqueta todos los objetos Tier‑0 con `Tag_Tier_Zero`; ChokeHound depende de esta clasificación.
   - En BloodHound CE -> Administration -> Configuration -> Early Access Features -> Enable Tier Management Engine​
   - Añade tus objectos Tier 0 en la sección "Privilege Zone Management". Tier 0 members to consider: [Tier Zero: Members and Modification](https://bloodhound.specterops.io/get-started/security-boundaries/tier-zero-members)


## Instalación

1. **Clonar el repositorio**  
   ```bash
   git clone https://github.com/ProtAAPP/ChokeHound.git
   cd ChokeHound
   ```
2. **Crear y activar un entorno virtual**
   - Windows (PowerShell)
     ```powershell
     python -m venv .venv
     .\.venv\Scripts\Activate.ps1
     ```
   - Linux / macOS
     ```bash
     python3 -m venv .venv
     source .venv/bin/activate
     ```
3. **Instalar dependencias**
   ```bash
   pip install -r requirements.txt
   ```
4. **Verificar Neo4j**  
   Asegúrate de que la instancia de Neo4j con datos de BloodHound esté en ejecución antes de lanzar el script.

## Configuración

Actualiza `config.py` con los valores de tu entorno (solo si es distinto que por defecto):

```python
NEO4J_URI = "bolt://localhost:7687"
NEO4J_USER = "neo4j"
NEO4J_PASSWORD = "bloodhoundcommunityedition"
DEFAULT_OUTPUT_FILENAME = "ChokeHound_report.xlsx"
LIMIT_CHOKE_POINTS = 200
```

## Uso

ChokeHound expone argumentos CLI para personalizar la salida.

## Uso Básico

```bash
python ChokeHound.py
```

Genera `ChokeHound_report.xlsx` con la configuración por defecto.

## Especificar Archivo de Salida

```bash
python ChokeHound.py --output mi_informe.xlsx
```

Crea un informe con nombre personalizado (el `.xlsx` se agrega automáticamente).

## Generar Log de Riesgo

```bash
python ChokeHound.py --output mi_informe.xlsx --log-risk
```

Produce el Excel y `mi_informe_risk_calculation_log.txt`, con el detalle de cada puntuación.

## Cálculo de Riesgo para Choke Points

El riesgo se calcula por punto de estrangulamiento usando:

```
Riesgo = (PesoObjetoOrigen × CategoríaObjetoOrigen) +
         (PesoTipoRelación × CategoríaTipoRelación) +
         (PesoObjetoDestino × CategoríaObjetoDestino) +
         (PesoRutasAfectadas × MultiplicadorRutas × 10)
```

El riesgo resultante se normaliza a un rango de 1–100 para facilitar la comparación.

_Nota: El modelo de cálculo de riesgo sigue en evolución. Ajusta los pesos y las categorías en `risk_config.py` según las características de cada entorno antes de confiar en las puntuaciones._


## Personalización

Personaliza pesos, categorías y límites según tus necesidades.

### Ajustar Pesos de Riesgo

Modifica `RISK_WEIGHTS` en `risk_config.py` para priorizar distintos componentes.

### Modificar Categorías de Riesgo

Edita `SOURCE_OBJECT_CATEGORIES`, `RELATIONSHIP_TYPE_CATEGORIES` y `TARGET_OBJECT_CATEGORIES` en `risk_config.py` para alinearlos a tu modelo de amenazas.

### Cambiar Límite de Resultados

Ajusta `LIMIT_CHOKE_POINTS` en `config.py` para controlar cuántos Choke Points priorizados se incluyen.

## TODO

- Añadir soporte para Azure / Entra ID.
- Analizar y estimar el riesgo de otros Choke Points que no impactan directamente Tier 0.
- Seguir mejorando y afinando la estimación de riesgo.
- Configurar otros Tiers (no solo Tier 0).
- Añadir recomendaciones para solucionar Choke Points detectados.

## Referencias

[BloodHound Community Edition Quickstart](https://bloodhound.specterops.io/get-started/quickstart/community-edition-quickstart)

[Tier Zero: Members and Modification](https://bloodhound.specterops.io/get-started/security-boundaries/tier-zero-members)

[Microsoft Enterprise Access Model](https://learn.microsoft.com/en-us/security/privileged-access-workstations/privileged-access-access-model)


## Autor

[@ciyinet](https://x.com/ciyinet)

[@gobispo](https://x.com/gobispo)

