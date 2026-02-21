"""
Tarifas ISR 2026 — Anexo 8 de la Resolución Miscelánea Fiscal 2026
Publicado en el DOF el 28 de diciembre de 2025.
Tarifas actualizadas por inflación acumulada >10% desde nov-2022 (Art. 152 LISR).

Incluye: tarifa mensual (Art. 96), tarifa anual (Art. 152),
subsidio al empleo, e INPC recientes.
"""

# ── Tarifa MENSUAL Art. 96 LISR (retenciones salarios) ────────────
TARIFA_MENSUAL = [
    {"lim_inf": 0.01,       "lim_sup": 844.59,      "cuota": 0.00,        "pct": 1.92},
    {"lim_inf": 844.60,     "lim_sup": 7_168.51,     "cuota": 16.22,       "pct": 6.40},
    {"lim_inf": 7_168.52,   "lim_sup": 12_598.02,    "cuota": 420.95,      "pct": 10.88},
    {"lim_inf": 12_598.03,  "lim_sup": 14_644.64,    "cuota": 1_011.68,    "pct": 16.00},
    {"lim_inf": 14_644.65,  "lim_sup": 17_533.64,    "cuota": 1_339.14,    "pct": 17.92},
    {"lim_inf": 17_533.65,  "lim_sup": 35_362.83,    "cuota": 1_856.84,    "pct": 21.36},
    {"lim_inf": 35_362.84,  "lim_sup": 55_736.68,    "cuota": 5_665.16,    "pct": 23.52},
    {"lim_inf": 55_736.69,  "lim_sup": 106_410.50,   "cuota": 10_457.09,   "pct": 30.00},
    {"lim_inf": 106_410.51, "lim_sup": 141_880.66,   "cuota": 25_659.23,   "pct": 32.00},
    {"lim_inf": 141_880.67, "lim_sup": 425_641.99,   "cuota": 37_009.69,   "pct": 34.00},
    {"lim_inf": 425_642.00, "lim_sup": 999_999_999,  "cuota": 133_488.54,  "pct": 35.00},
]

# ── Tarifa ANUAL Art. 152 LISR (cálculo del ejercicio) ────────────
TARIFA_ANUAL = [
    {"lim_inf": 0.01,          "lim_sup": 10_135.11,      "cuota": 0.00,         "pct": 1.92},
    {"lim_inf": 10_135.12,     "lim_sup": 86_022.11,      "cuota": 194.59,       "pct": 6.40},
    {"lim_inf": 86_022.12,     "lim_sup": 151_176.19,     "cuota": 5_051.37,     "pct": 10.88},
    {"lim_inf": 151_176.20,    "lim_sup": 175_735.66,     "cuota": 12_140.13,    "pct": 16.00},
    {"lim_inf": 175_735.67,    "lim_sup": 210_403.69,     "cuota": 16_069.64,    "pct": 17.92},
    {"lim_inf": 210_403.70,    "lim_sup": 424_353.97,     "cuota": 22_282.14,    "pct": 21.36},
    {"lim_inf": 424_353.98,    "lim_sup": 668_840.14,     "cuota": 67_981.92,    "pct": 23.52},
    {"lim_inf": 668_840.15,    "lim_sup": 1_276_925.98,   "cuota": 125_485.07,   "pct": 30.00},
    {"lim_inf": 1_276_925.99,  "lim_sup": 1_702_567.97,   "cuota": 307_910.81,   "pct": 32.00},
    {"lim_inf": 1_702_567.98,  "lim_sup": 5_107_703.92,   "cuota": 444_116.23,   "pct": 34.00},
    {"lim_inf": 5_107_703.93,  "lim_sup": 999_999_999_99, "cuota": 1_601_862.46, "pct": 35.00},
]

# ── Subsidio para el Empleo (tabla mensual) ───────────────────────
SUBSIDIO_EMPLEO_MENSUAL = [
    {"desde": 0.01,     "hasta": 1_768.96,  "subsidio": 407.02},
    {"desde": 1_768.97, "hasta": 2_653.38,  "subsidio": 406.83},
    {"desde": 2_653.39, "hasta": 3_472.84,  "subsidio": 406.62},
    {"desde": 3_472.85, "hasta": 3_537.87,  "subsidio": 392.77},
    {"desde": 3_537.88, "hasta": 4_446.15,  "subsidio": 382.46},
    {"desde": 4_446.16, "hasta": 4_717.18,  "subsidio": 354.23},
    {"desde": 4_717.19, "hasta": 5_335.42,  "subsidio": 324.87},
    {"desde": 5_335.43, "hasta": 6_224.67,  "subsidio": 294.63},
    {"desde": 6_224.68, "hasta": 7_113.90,  "subsidio": 253.54},
    {"desde": 7_113.91, "hasta": 7_382.33,  "subsidio": 217.61},
    {"desde": 7_382.34, "hasta": 999_999,   "subsidio": 0.00},
]

# ── INPC (valores recientes para Factor de Actualización) ─────────
# Fuente: INEGI. Últimos valores disponibles.
INPC = {
    "2022-11": 122.515,   # Nov 2022 (base para actualización)
    "2023-12": 131.507,
    "2024-06": 133.271,
    "2024-12": 136.163,
    "2025-06": 138.500,   # Estimado
    "2025-12": 141.200,   # Estimado
}

# INPC para ejercicios de Factor de Actualización
INPC_RECIENTE = 141.200    # Dic 2025 (estimado)
INPC_ANTERIOR = 136.163    # Dic 2024

# ── Funciones auxiliares ──────────────────────────────────────────

def calcular_isr_mensual(base_gravable: float) -> dict:
    """Calcula ISR mensual Art. 96 LISR dado una base gravable."""
    for rango in TARIFA_MENSUAL:
        if rango["lim_inf"] <= base_gravable <= rango["lim_sup"]:
            excedente = base_gravable - rango["lim_inf"]
            isr_marginal = excedente * (rango["pct"] / 100)
            isr_total = isr_marginal + rango["cuota"]
            return {
                "base_gravable": base_gravable,
                "lim_inf": rango["lim_inf"],
                "excedente": excedente,
                "pct": rango["pct"],
                "isr_marginal": round(isr_marginal, 2),
                "cuota_fija": rango["cuota"],
                "isr_total": round(isr_total, 2),
            }
    return {}


def calcular_isr_anual(base_gravable: float) -> dict:
    """Calcula ISR anual Art. 152 LISR dado una base gravable."""
    for rango in TARIFA_ANUAL:
        if rango["lim_inf"] <= base_gravable <= rango["lim_sup"]:
            excedente = base_gravable - rango["lim_inf"]
            isr_marginal = excedente * (rango["pct"] / 100)
            isr_total = isr_marginal + rango["cuota"]
            return {
                "base_gravable": base_gravable,
                "lim_inf": rango["lim_inf"],
                "excedente": excedente,
                "pct": rango["pct"],
                "isr_marginal": round(isr_marginal, 2),
                "cuota_fija": rango["cuota"],
                "isr_total": round(isr_total, 2),
            }
    return {}
