const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");

const app = express();
const upload = multer({ dest: "uploads/" });

// Middleware für JSON- und URL-codierte Daten
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Route: Excel-Datei hochladen und verarbeiten
app.post("/api/excel", upload.single("file"), (req, res) => {
    try {
        const filePath = req.file.path;
        const newData = req.body.newData;

        // Validierung: Überprüfen, ob die neuen Daten als Array übergeben wurden
        if (!newData || !Array.isArray(JSON.parse(newData))) {
            return res.status(400).json({ error: "newData muss ein Array sein, z.B.: [\"Wert1\", \"Wert2\", \"Wert3\"]" });
        }

        const newRow = JSON.parse(newData); // Konvertiere die Daten aus dem JSON-String in ein Array

        // Excel-Datei einlesen
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Bereich der aktuellen Tabelle ermitteln
        const range = xlsx.utils.decode_range(worksheet["!ref"]);
        const nextRow = range.e.r + 1; // Die nächste freie Zeile

        // Die neuen Werte in die nächste Zeile einfügen
        newRow.forEach((value, index) => {
            const cellAddress = `${xlsx.utils.encode_col(index)}${nextRow + 1}`; // Spaltenindex -> Buchstabe
            worksheet[cellAddress] = { v: value };
        });

        // Bereich aktualisieren
        range.e.r++;
        worksheet["!ref"] = xlsx.utils.encode_range(range);

        // Aktualisierte Datei als Buffer speichern
        const updatedFileBuffer = xlsx.write(workbook, { type: "buffer", bookType: "xlsx" });

        // Temporäre Datei löschen
        fs.unlinkSync(filePath);

        // Datei als Download zurückgeben
        res.setHeader("Content-Disposition", "attachment; filename=updated_file.xlsx");
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.send(updatedFileBuffer);
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: "Fehler beim Verarbeiten der Datei." });
    }
});

// Root-Route
app.get("/", (req, res) => {
    res.json({ message: "Willkommen bei der Excel-API!" });
});

app.post("/convert-to-csv", upload.single("file"), (req, res) => {
    try {
        const filePath = req.file.path;
        const newData = req.body.newData;

        // Validierung: Überprüfen, ob die neuen Daten als Array übergeben wurden
        if (!newData || !Array.isArray(JSON.parse(newData))) {
            return res.status(400).json({ error: "newData muss ein Array sein, z.B.: [\"Wert1\", \"Wert2\", \"Wert3\"]" });
        }

        const newRow = JSON.parse(newData); // Konvertiere die Daten aus dem JSON-String in ein Array

        // Excel-Datei einlesen
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const data = xlsx.utils.sheet_to_json(worksheet, {header: 1});
        data.push(newRow);

        const updatedFilePath = `uploads/updated_${req.file.originalname}`;
        xlsx.writeFile(workbook, updatedFilePath);

        // Temporäre Datei löschen
        fs.unlinkSync(filePath);

        // Datei als Download zurückgeben
        res.setHeader("Content-Disposition", "attachment; filename=updated_file.xlsx");
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.sendFile(updatedFilePath, (err) => {
            if (err) {
                console.error("Fehler beim Senden der Datei:",err);
                res.status(500).send("Fehler beim Senden der Datei.");
            }
        })
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: "Fehler beim Verarbeiten der Datei." });
    }
});

// Server starten
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server läuft auf http://localhost:${PORT}`);
});
