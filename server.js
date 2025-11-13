// server.js
import express from "express";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import ExcelJS from "exceljs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;
const XLSX_PATH = path.join(__dirname, "reponses.xlsx");

app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// Colonnes de base (ORDRE dans le fichier)
// temps_reponse avant prénom, ajout de "score"
const BASE_HEADERS = [
  "temps_reponse",
  "score",
  "prenom",
  "nom",
  "etablissement",
  "poste",
  "date"
];

// Création / récupération du classeur
async function getWorkbook() {
  const workbook = new ExcelJS.Workbook();

  if (fs.existsSync(XLSX_PATH)) {
    await workbook.xlsx.readFile(XLSX_PATH);
  }

  let sheet = workbook.getWorksheet("Réponses");
  if (!sheet) {
    sheet = workbook.addWorksheet("Réponses");

    // Ligne 1 : labels visibles (au début = clés de base)
    const labelRow = sheet.addRow(BASE_HEADERS);

    // Ligne 2 : clés techniques (toujours)
    const keyRow = sheet.addRow(BASE_HEADERS);
    keyRow.hidden = true;

    // Style en-tête
    labelRow.font = { bold: true };
    labelRow.alignment = { vertical: "middle", wrapText: true };
    BASE_HEADERS.forEach((_, idx) => {
      sheet.getColumn(idx + 1).width = 22;
    });

    // Style spécifique colonne "score" (header)
    const scoreIndex = BASE_HEADERS.indexOf("score") + 1; // 1-based
    if (scoreIndex > 0) {
      const scoreHeaderCell = labelRow.getCell(scoreIndex);
      scoreHeaderCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFDE68A" } // jaune pâle
      };
      scoreHeaderCell.font = { bold: true, color: { argb: "FF854D0E" } };
    }
  } else {
    // Si le fichier n'existait pas, le code ci-dessus a déjà créé la feuille
    await workbook.xlsx.writeFile(XLSX_PATH);
  }

  // Migration éventuelle si ancien fichier sans ligne 2 (clés)
  sheet = workbook.getWorksheet("Réponses");
  if (sheet.rowCount === 1) {
    const labelRow = sheet.getRow(1);
    const vals = labelRow.values.slice(1).map(v => String(v || ""));
    const keyRow = sheet.addRow(vals);
    keyRow.hidden = true;
  }

  return workbook;
}

app.post("/api/save", async (req, res) => {
  try {
    const d = req.body || {};

    const workbook = await getWorkbook();
    const sheet = workbook.getWorksheet("Réponses");

    const labelRow = sheet.getRow(1);
    const keyRow = sheet.getRow(2);

    // Clés existantes (ligne 2)
    let keys = keyRow.values.slice(1).map(v => String(v || ""));

    // Si jamais le fichier était ancien / vide, on s'assure que BASE_HEADERS sont là
    if (keys.length === 0) {
      keys = [...BASE_HEADERS];
      keyRow.values = [ , ...keys ];
      keyRow.hidden = true;
    } else {
      // S'assurer que toutes les colonnes de base existent au début
      BASE_HEADERS.forEach(baseKey => {
        if (!keys.includes(baseKey)) {
          keys.unshift(baseKey);
        }
      });
      keys = Array.from(new Set(keys)); // dédoublonnage
      keyRow.values = [ , ...keys ];
      keyRow.hidden = true;
    }

    // --- Colonnes de questions dynamiques ---
    const labelsMap = d.__labels || {};
    const choiceKeys = Object.keys(d).filter(k => k.endsWith("_choice"));

    // On ajoute les nouvelles clés de questions à la liste des keys si pas présentes
    for (const key of choiceKeys) {
      if (!keys.includes(key)) {
        keys.push(key);
      }
    }

    // recalcul des values de la ligne 2 (clé techniques)
    keyRow.values = [ , ...keys ];
    keyRow.hidden = true;

    // --- Calcul du score (nombre de A) ---
    const correctCount = choiceKeys.filter(k => d[k] === "A").length;
    d.score = correctCount;

    // --- Ligne 1 : labels visibles ---
    const visibleHeaders = keys.map((k) => {
      // Pour les colonnes de base, on garde le même texte
      if (BASE_HEADERS.includes(k)) return k;

      // Pour les questions *_choice, on affiche l'intitulé
      return labelsMap[k] || k;
    });

    labelRow.values = [ , ...visibleHeaders ];
    labelRow.font = { bold: true };
    labelRow.alignment = { vertical: "middle", wrapText: true };

    // Largeur des colonnes selon le texte
    visibleHeaders.forEach((txt, i) => {
      const len = String(txt || "").length;
      sheet.getColumn(i + 1).width = Math.min(Math.max(len * 0.9, 22), 60);
    });

    // Style aussi la colonne "score" au niveau du header
    const scoreIndex = keys.indexOf("score") + 1; // 1-based
    if (scoreIndex > 0) {
      const scoreHeaderCell = labelRow.getCell(scoreIndex);
      scoreHeaderCell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFDE68A" } // jaune clair
      };
      scoreHeaderCell.font = { bold: true, color: { argb: "FF854D0E" } };
    }

    // --- Ajout de la ligne de données ---
    const rowValues = keys.map(h => d[h] ?? "");
    const newRow = sheet.addRow(rowValues);

    // Style wrap + couleur pour A / autres + style colonne score
    for (let col = 1; col <= newRow.cellCount; col++) {
      const cell = newRow.getCell(col);
      cell.alignment = { wrapText: true, vertical: "top" };

      const key = keys[col - 1] || "";

      // cellule de score → fond jaune/doré
      if (key === "score") {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFDE68A" } // jaune doux
        };
        cell.font = { bold: true, color: { argb: "FF854D0E" } };
      }

      // colonnes de réponses QCM
      if (key.endsWith("_choice")) {
        const val = d[key]; // "A"/"B"/"C"/""
        if (val === "A") {
          // vert doux pour bonne réponse
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFE8F7EE" }
          };
          cell.font = { color: { argb: "FF166534" } };
          cell.border = {
            top:    { style: "thin", color: { argb: "FF16A34A" } },
            left:   { style: "thin", color: { argb: "FF16A34A" } },
            bottom: { style: "thin", color: { argb: "FF16A34A" } },
            right:  { style: "thin", color: { argb: "FF16A34A" } }
          };
        } else if (val && val !== "A") {
          // rouge doux pour mauvaise réponse
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFEE2E2" }
          };
          cell.font = { color: { argb: "FF7F1D1D" } };
          cell.border = {
            top:    { style: "thin", color: { argb: "FFEF4444" } },
            left:   { style: "thin", color: { argb: "FFEF4444" } },
            bottom: { style: "thin", color: { argb: "FFEF4444" } },
            right:  { style: "thin", color: { argb: "FFEF4444" } }
          };
        }
      }
    }

    await workbook.xlsx.writeFile(XLSX_PATH);
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: "xlsx_write_failed" });
  }
});

app.get("/export-xlsx", async (req, res) => {
  if (!fs.existsSync(XLSX_PATH)) {
    return res.status(404).send("Pas encore de données.");
  }
  res.download(XLSX_PATH);
});

app.listen(PORT, () => {
  console.log(`🚀 Serveur lancé sur http://localhost:${PORT}`);
});
