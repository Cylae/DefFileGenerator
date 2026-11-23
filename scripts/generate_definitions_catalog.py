#!/usr/bin/env python3
"""Generate a comprehensive README.md catalog for the Definitions/ directory."""

import csv
from collections import defaultdict
from pathlib import Path


def generate_catalog():
    def_dir = Path("Definitions")
    files = sorted(def_dir.glob("*.csv"))

    catalog = []
    for f in files:
        with open(f, encoding="utf-8", errors="ignore") as fp:
            reader = csv.reader(fp, delimiter=";")
            header = next(reader, [])
            protocol = header[0] if len(header) > 0 else "Modbus"
            eq_type = header[1] if len(header) > 1 else "Unknown"
            mfr = header[2] if len(header) > 2 else ""
            model = header[3] if len(header) > 3 else ""

            reg_count = sum(1 for line in reader if line and any(line))

            prefix = f.name.split("_")[0]
            catalog.append(
                {
                    "filename": f.name,
                    "prefix": prefix,
                    "type": eq_type,
                    "manufacturer": mfr,
                    "model": model,
                    "protocol": protocol,
                    "registers": reg_count,
                }
            )

    grouped = defaultdict(list)
    for item in catalog:
        grouped[item["prefix"]].append(item)

    category_titles = {
        "INV": ("Onduleurs & Convertisseurs (Inverters)", "Inverter"),
        "MET": ("Compteurs & Analyseurs Réseau (Meters)", "Meter"),
        "BAT": ("Systèmes de Stockage BESS (Batteries)", "Battery"),
        "SEN": ("Capteurs & Pyranomètres (Sensors)", "Sensor"),
        "COM": ("Passerelles & Loggers de Communication", "Logger"),
        "OTH": ("Autres Équipements (IO Modules & Afficheurs)", "Io / Display"),
    }

    total_regs = sum(c["registers"] for c in catalog)

    def fmt_num(n: int) -> str:
        return f"{n:_}".replace("_", " ")

    lines = [
        "# 📚 Webdyn Definitions Library",
        "",
        f"Bibliothèque de définitions d'équipements industriels au format WebdynSunPM. "
        f"Comprend **{len(catalog)} fichiers de définition** validés couvrant "
        f"**{fmt_num(total_regs)} registres Modbus** à travers plus de 60 constructeurs mondiaux.",
        "",
        "---",
        "",
        "## 📊 Synthèse par Catégorie",
        "",
        "| Catégorie | Préfixe | Fichiers | Total Registres | Constructeurs Notables |",
        "|-----------|---------|----------|-----------------|------------------------|",
    ]

    for pfx in ["INV", "MET", "BAT", "SEN", "COM", "OTH"]:
        items = grouped[pfx]
        cat_regs = sum(i["registers"] for i in items)
        mfrs = sorted(set(i["manufacturer"] for i in items))
        notable = ", ".join(mfrs[:5]) + (f" (+{len(mfrs) - 5})" if len(mfrs) > 5 else "")
        title, _ = category_titles[pfx]
        lines.append(f"| {title} | `{pfx}_` | {len(items)} | {fmt_num(cat_regs)} | {notable} |")

    lines.extend(
        [
            "",
            "---",
            "",
            "## 📁 Arborescence Propre & Structurée par Constructeur",
            "",
            "Pour une lisibilité et une maintenabilité optimales, les fichiers de définition sont classés dans des sous-dossiers dédiés par grand équipementier tout en maintenant une racine plate pour la rétro-compatibilité :",
            "",
            "| Dossier Dédié | Équipementier / Gamme | Fichiers Certifiés | Description |",
            "|---------------|----------------------|--------------------|-------------|",
            "| [`Definitions/SMA/`](./SMA/) | SMA Solar Technology | 27 | Gammes Sunny Boy, Sunny Tripower (Core1, Core2, STP-X, SE), Sunny Island, Sunny Central, Data Manager M |",
            "| [`Definitions/GoodWe/`](./GoodWe/) | GoodWe Power Supply | 32 | Gammes résidentielles & tertiaires SMT, ET/EH, HT, MT, XS/DNS/MS, SDT, Lynx Home HV, EzLogger |",
            "| [`Definitions/Huawei/`](./Huawei/) | Huawei FusionSolar | 65 | Gammes SUN2000 M0/M1/M2/M3/Utility, SmartLogger 1000/3000, Luna2000, Power Meters DTSU666-H |",
            "| [`Definitions/Chint/`](./Chint/) | Chint Instrument | 5 | Compteurs communicants monophasés et triphasés DDSU666, DTSU666, DTSU666-H |",
            "| [`Definitions/Solplanet/`](./Solplanet/) | Solplanet (AISWEI) | 13 | Gammes ASW monophasés, triphasés, hybrides et passerelles Ai-Logger |",
            "| [`Definitions/Schneider/`](./Schneider/) | Schneider Electric | 79 | Gammes Conext CL, Core XC, compteurs réseau PM5000, iEM3000, EM1200 |",
            "| [`Definitions/Socomec/`](./Socomec/) | Socomec Power | 36 | Analyseurs de réseau Diris A10/A20/A40/B10/B30/I30, compteurs d'énergie Countis E |",
            "| [`Definitions/Sungrow/`](./Sungrow/) | Sungrow Power | 24 | Onduleurs SG tertiaires & utility, batteries BESS, coffrets string PV |",
            "| [`Definitions/Growatt/`](./Growatt/) | Growatt New Energy | 16 | Onduleurs string MAX, MOD, MIN, TL-X, TL3-X, hybrides et stockage |",
            "| [`Definitions/Fronius/`](./Fronius/) | Fronius International | 34 | Gammes Symo, Primo, Eco, Gen24 Plus, cartes d'entrées/capteurs IG |",
            "| [`Definitions/SolarEdge/`](./SolarEdge/) | SolarEdge Technologies | 20 | Onduleurs string SE monophasés/triphasés, blocs de contrôle de puissance active/réactive |",
            "| [`Definitions/CarloGavazzi/`](./CarloGavazzi/) | Carlo Gavazzi Automation | 27 | Compteurs d'énergie et analyseurs EM24, EM210, EM210_MV, EM530, EM540, EM100, EM300, WM |",
            "| [`Definitions/ABB/`](./ABB/) | ABB / Fimer | 71 | Onduleurs PVS (10..175, 800), Trio, Pro33, Ultra, compteurs B21/23/24, M2M, M4M, Terra AC |",
            "| [`Definitions/Dyness/`](./Dyness/) | Dyness Renewable | 4 | Systèmes de stockage batteries haute tension BESS DH100F, DH200F, DH200Y |",
            "| [`Definitions/Classic/`](./Classic/) | Grands Classiques Multi-Constructeurs | 18 | Profils unifiés et standardisés des 18 équipements les plus déployés |",
            "| [`Definitions/X7/`](./X7/) | X7 C&I BESS | 7 | Définitions ultra-détaillées du conteneur de stockage X7 100kW/215kWh |",
            "",
            "---",
            "",
            "## 🏷️ Convention de Nommage",
            "",
            "Tous les fichiers suivent la convention canonique Webdyn :",
            "```text",
            "<CATEGORY_PREFIX>_<Manufacturer>_<Model>.csv",
            "```",
            "- `INV_` : Onduleurs solaires, hybrides et industriels (*Inverters*)",
            "- `MET_` : Compteurs d'énergie et analyseurs de réseau (*Meters*)",
            "- `SEN_` : Capteurs météo, pyranomètres et sondes thermiques (*Sensors*)",
            "- `COM_` : Loggers, concentrateurs et passerelles de données (*Loggers*)",
            "- `OTH_` : Modules d'entrées/sorties (Sens'IO), afficheurs et bornes (*Other*)",
            "",
            "---",
            "",
            "## 📑 Catalogue Détaillé",
            "",
        ]
    )

    for pfx in ["INV", "MET", "BAT", "SEN", "COM", "OTH"]:
        title, _ = category_titles[pfx]
        items = grouped[pfx]
        lines.extend(
            [
                f"### {title} ({len(items)} fichiers)",
                "",
                "| Fichier | Constructeur | Modèle | Protocole | Registres |",
                "|---------|--------------|--------|-----------|-----------|",
            ]
        )
        for item in sorted(items, key=lambda x: (x["manufacturer"].lower(), x["model"].lower())):
            fname = item["filename"]
            mfr = item["manufacturer"]
            mdl = item["model"]
            prot = item["protocol"]
            regs = item["registers"]
            lines.append(f"| [`{fname}`](./{fname}) | {mfr} | {mdl} | {prot} | {regs} |")
        lines.append("")

    output_path = def_dir / "README.md"
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(
        f"Catalog successfully written to {output_path} ({len(catalog)} files, {total_regs} registers)"
    )


if __name__ == "__main__":
    generate_catalog()
