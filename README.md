# ChokeHound - Tier 0 Choke Points Analysis Tool

[English](#chokehound---tier-0-choke-points-analysis-tool) | [Español](#chokehound---herramienta-de-análisis-de-choke-points-tier-0)

<p align="center">
  <img src="logo.png" alt="ChokeHound logo" width="240">
</p>

ChokeHound is a specialized tool to identify and analyze Tier 0 choke points in Active Directory and Azure/Entra ID environments from data collected by BloodHound. It analyzes BloodHound CE data to identify critical privilege edges connecting non-Tier-0 to Tier-0 objects. The generated Excel report helps defenders prioritize remediation by focusing on the most critical access paths.

## What are Choke Points?

Choke points are critical privilege edges that connect non-Tier-0 objects to Tier-0 objects in Active Directory. These represent the most dangerous access paths that attackers can exploit to escalate privileges and compromise your domain. By identifying and remediating these choke points, you can significantly reduce your attack surface.

## Features

- **Tier 0 Choke Point Analysis**: Identify critical privilege edges connecting non-Tier-0 to Tier-0 objects
- **Risk Scoring**: Prioritize findings using weighted risk scores based on multiple factors
- **Attack Path Analysis**: Understand how many attack paths are affected by each choke point
- **Excel Reports**: Professional Excel reports with formatting, color-coding, and detailed documentation
- **Relationship Documentation**: Automatic links to BloodHound documentation for each relationship type

## Prerequisites

- Python 3.7 or later
- Neo4j 4.x+ accessible to the script
- BloodHound CE data already ingested into Neo4j
- Network access to the BloodHound Neo4j instance

## Requirements Before Execution

The following steps are required for ChokeHound to produce meaningful results:

1. **BloodHound CE installation**: Deploy BloodHound CE and confirm the Neo4j backend is initialized and reachable with valid credentials. [Instructions here](https://bloodhound.specterops.io/get-started/quickstart/community-edition-quickstart).
2. **Data collection with SharpHound or AzureHound**: Run the official collectors and import the output into BloodHound CE: [SharpHound](https://bloodhound.specterops.io/collect-data/ce-collection/sharphound) on-prem / [AzureHound](https://bloodhound.specterops.io/collect-data/ce-collection/azurehound) for Entra ID.
3. **Tier 0 configuration inside BloodHound**: Tag every Tier‑0 object with `Tag_Tier_Zero` so ChokeHound can differentiate privileged targets.
   - In BloodHound CE -> Administration -> Configuration -> Early Access Features -> Enable Tier Management Engine
   - Add your Tier 0 objects in section "Privilege Zone Management". Tier 0 members to consider: [Tier Zero: Members and Modification](https://bloodhound.specterops.io/get-started/security-boundaries/tier-zero-members)

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

Adjust `chokehound/config/settings.py` to match your Neo4j endpoint (only if different from default):

```python
NEO4J_URI = "bolt://localhost:7687"
NEO4J_USER = "neo4j"
NEO4J_PASSWORD = "bloodhoundcommunityedition"
DEFAULT_OUTPUT_FILENAME = "ChokeHound_report.xlsx"
```

## Usage

ChokeHound uses CLI arguments for output customization and logging.

### Basic Usage

```bash
python chokehound.py
```

Generates `ChokeHound_report.xlsx` using default settings.

### Specify Output File

```bash
python chokehound.py --output my_report.xlsx
```

Creates a report with a custom filename (the `.xlsx` extension is appended automatically).

### Generate Risk Calculation Log

```bash
python chokehound.py --output my_report.xlsx --log-risk
```

Produces both the Excel workbook and `my_report_risk_calculation_log.txt`, which documents each choke point's scoring breakdown.

### Skip Risk Calculation (Fast Mode)

```bash
python chokehound.py --skip-risk-calculation
```

Runs the analysis without calculating risk scores. This provides faster execution and simpler output, with results sorted by the number of affected attack paths instead of risk score. Useful for quick scans or when you only need to identify choke points without prioritization.

## Architecture

ChokeHound is built with a modular architecture:

```
chokehound/
├── core/              # Core functionality
│   ├── database.py    # Neo4j connection management
│   └── query_executor.py  # Query execution engine
├── queries/           # Choke points query module
│   ├── registry.py    # Query registration system
│   └── choke_points.py  # Choke point analysis queries
├── reporting/         # Report generation
│   └── excel_report.py  # Excel report generator
├── config/            # Configuration
│   ├── settings.py    # Application settings
│   └── risk_config.py # Risk scoring configuration
└── utils/             # Utility functions
    └── label_processor.py  # Label processing utilities
```

## Risk Calculation for Choke Points

ChokeHound calculates a weighted risk score per choke point:

```
Risk = (SourceObjectWeight × SourceObjectCategory) +
       (RelationshipTypeWeight × RelationshipTypeCategory) +
       (TargetObjectWeight × TargetObjectCategory) +
       (AffectedAttackPathsWeight × PathsMultiplier × 10)
```

The resulting risk is normalized to a 1–100 scale, ensuring consistent comparison across findings.

_Note: The scoring model is still evolving. Review and adapt the weights and category values in `chokehound/config/risk_config.py` for your environment before relying on the scores._

## Customization

### Adjust Risk Weights

Update the `RISK_WEIGHTS` dictionary in `chokehound/config/risk_config.py` to emphasize or de-emphasize specific components.

### Modify Risk Categories

Edit `SOURCE_OBJECT_CATEGORIES`, `RELATIONSHIP_TYPE_CATEGORIES`, and `TARGET_OBJECT_CATEGORIES` in `chokehound/config/risk_config.py` to reflect your threat model.

### Change Hop Limit and Risk Weights

All Tier 0 choke points configuration is located in `chokehound/config/settings.py`:
- `AD_CHOKE_POINTS_HOP_LIMIT`: Adjusts how deep the script traverses upstream relationships when counting affected attack paths (default: 3)
- `AD_CHOKE_POINTS_LIMIT`: Maximum number of AD choke points to return (default: 200)
- `AZURE_CHOKE_POINTS_HOP_LIMIT`: Hop limit for Azure queries (default: 3)
- `AZURE_CHOKE_POINTS_LIMIT`: Maximum number of Azure choke points to return (default: 200)

Risk scoring configuration is located in `chokehound/config/risk_config.py`:
- `RISK_WEIGHTS`: Adjusts the weight of each component in the risk calculation
- Risk categories and SID patterns for source/target objects and relationships

## Included Analysis

- **AD Tier 0 Choke Points Risk**: Prioritized Active Directory Tier 0 choke points connecting non-Tier-0 to Tier-0 objects with risk scoring
- **Azure Tier 0 Choke Points**: Azure/Entra ID choke points connecting non-Tier-0 to Tier-0 objects
- **ADirect relationships into Tier0**: Ranking of Active Directory relationship types that connect to Tier-0
- **Source nodes into T0 Rank**: Top non-Tier-0 nodes ranked by number of Tier-0 targets they can reach

## References

- [BloodHound Community Edition Quickstart](https://bloodhound.specterops.io/get-started/quickstart/community-edition-quickstart)
- [Tier Zero: Members and Modification](https://bloodhound.specterops.io/get-started/security-boundaries/tier-zero-members)
- [Microsoft Enterprise Access Model](https://learn.microsoft.com/en-us/security/privileged-access-workstations/privileged-access-access-model)

## Authors

- [@ciyinet](https://github.com/ciyinet)
- [@gobispo](https://github.com/gobispo)

---

# ChokeHound - Herramienta de Análisis de Choke Points Tier 0

[English](#chokehound---tier-0-choke-points-analysis-tool) | [Español](#chokehound---herramienta-de-análisis-de-choke-points-tier-0)

ChokeHound es una herramienta especializada para identificar y analizar choke points de Tier 0 en entornos de Active Directory y Azure/Entra ID a partir de datos recopilados por BloodHound. Analiza datos de BloodHound CE para identificar relaciones de privilegio críticas que conectan objetos no Tier-0 con objetos Tier-0. El informe Excel generado ayuda a los defensores a priorizar la mitigación enfocándose en las rutas de acceso más críticas.

## ¿Qué son los Choke Points?

Los choke points son relaciones de privilegio críticas que conectan objetos no Tier-0 con objetos Tier-0 en Active Directory. Representan las rutas de acceso más peligrosas que los atacantes pueden explotar para escalar privilegios y comprometer tu dominio. Al identificar y remediar estos choke points, puedes reducir significativamente tu superficie de ataque.

## Características

- **Análisis de Choke Points de Tier 0**: Identifica relaciones de privilegio críticas que conectan objetos no Tier-0 con Tier-0
- **Puntuación de Riesgo**: Prioriza hallazgos usando puntuaciones de riesgo ponderadas basadas en múltiples factores
- **Análisis de Rutas de Ataque**: Entiende cuántas rutas de ataque son afectadas por cada choke point
- **Informes Excel**: Informes Excel profesionales con formato, codificación de colores y documentación detallada
- **Documentación de Relaciones**: Enlaces automáticos a la documentación de BloodHound para cada tipo de relación

## Requisitos Previos

- Python 3.7 o superior
- Neo4j 4.x+ accesible para el script
- Datos de BloodHound CE ya importados en Neo4j
- Acceso de red a la instancia de Neo4j de BloodHound

## Requisitos Antes de la Ejecución

Los siguientes pasos son necesarios para que ChokeHound produzca resultados significativos:

1. **Instalación de BloodHound CE**: Despliega BloodHound CE y confirma que Neo4j está inicializado y accesible con credenciales válidas. [Instrucciones aquí](https://bloodhound.specterops.io/get-started/quickstart/community-edition-quickstart).
2. **Recolección de datos con SharpHound o AzureHound**: Ejecuta los colectores oficiales e importa los resultados en BloodHound CE: [SharpHound](https://bloodhound.specterops.io/collect-data/ce-collection/sharphound) on-prem / [AzureHound](https://bloodhound.specterops.io/collect-data/ce-collection/azurehound) para Entra ID.
3. **Configuración de Tier 0 en BloodHound**: Etiqueta todos los objetos Tier‑0 con `Tag_Tier_Zero` para que ChokeHound pueda diferenciar objetivos privilegiados.
   - En BloodHound CE -> Administration -> Configuration -> Early Access Features -> Enable Tier Management Engine
   - Añade tus objetos Tier 0 en la sección "Privilege Zone Management". Tier 0 members to consider: [Tier Zero: Members and Modification](https://bloodhound.specterops.io/get-started/security-boundaries/tier-zero-members)

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

Actualiza `chokehound/config/settings.py` con los valores de tu entorno (solo si es distinto que por defecto):

```python
NEO4J_URI = "bolt://localhost:7687"
NEO4J_USER = "neo4j"
NEO4J_PASSWORD = "bloodhoundcommunityedition"
DEFAULT_OUTPUT_FILENAME = "ChokeHound_report.xlsx"
```

## Uso

ChokeHound expone argumentos CLI para personalizar la salida.

### Uso Básico

```bash
python chokehound.py
```

Genera `ChokeHound_report.xlsx` con la configuración por defecto.

### Especificar Archivo de Salida

```bash
python chokehound.py --output mi_informe.xlsx
```

Crea un informe con nombre personalizado (el `.xlsx` se agrega automáticamente).

### Generar Log de Riesgo

```bash
python chokehound.py --output mi_informe.xlsx --log-risk
```

Produce el Excel y `mi_informe_risk_calculation_log.txt`, con el detalle de cada puntuación.

### Omitir Cálculo de Riesgo (Modo Rápido)

```bash
python chokehound.py --skip-risk-calculation
```

Ejecuta el análisis sin calcular puntuaciones de riesgo. Esto proporciona una ejecución más rápida y resultados más simples, ordenados por el número de rutas de ataque afectadas en lugar de puntuación de riesgo. Útil para análisis rápidos o cuando solo necesitas identificar choke points sin priorización.

## Arquitectura

ChokeHound está construido con una arquitectura modular. Ver la sección en inglés para más detalles.

## Cálculo de Riesgo para Choke Points

El riesgo se calcula por choke point usando la fórmula descrita en la sección en inglés.

## Personalización

Ver la sección en inglés para detalles sobre personalización de pesos de riesgo, categorías y límites.

## Análisis Incluido

- **AD Tier 0 Choke Points Risk**: Choke points de Tier 0 de Active Directory priorizados que conectan objetos no Tier-0 con Tier-0 con puntuación de riesgo
- **Azure Tier 0 Choke Points**: Choke points de Azure/Entra ID que conectan objetos no Tier-0 con Tier-0
- **ADirect relationships into Tier0**: Ranking de tipos de relaciones de Active Directory que conectan a Tier-0
- **Source nodes into T0 Rank**: Top nodos no Tier-0 clasificados por número de objetivos Tier-0 que pueden alcanzar

## Referencias

- [BloodHound Community Edition Quickstart](https://bloodhound.specterops.io/get-started/quickstart/community-edition-quickstart)
- [Tier Zero: Members and Modification](https://bloodhound.specterops.io/get-started/security-boundaries/tier-zero-members)
- [Microsoft Enterprise Access Model](https://learn.microsoft.com/en-us/security/privileged-access-workstations/privileged-access-access-model)
